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
                                                  DK # Danmark
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

================================={ Example Common LDAP Filters (for reference/docs) }=================================

Common LDAP Filters:

All users:                          (objectClass=user)
All enabled users:                  (&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))
All disabled users:                 (&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))
All groups:                         (objectClass=group)
All computers:                      (objectClass=computer)
Users with email:                   (&(objectClass=user)(mail=*))
Users in specific OU (need DN):     (&(objectClass=user)(distinguishedName=*,OU=IT,DC=example,DC=com))
Users created after date:           (&(objectClass=user)(whenCreated>=20240101000000.0Z))
Users whose password never expires: (&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))
Users with admin in name:           (&(objectClass=user)(name=*admin*))
Security groups:                    (&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=2147483648))
Distribution groups:                (&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))

DEVICE NAMING & DEMO DATA CONVENTIONS
=====================================

Company Suffixes:

Name                                 Country             Legal Form            Notes
--------------------------------     ------------------ --------------------  ------------------------
Example Music (England) Ltd          England & Wales     Ltd                  Companies House jurisdiction
Example Music (Scotland) Ltd         Scotland            Ltd (SCxxxxxx)       Scottish company numbers start with "SC"
Example Music (Danmark) ApS          Denmark             ApS                  Danmark Aktieselskab (public limited)
Example Music (Deutschland) GmbH     Germany             GmbH                 Gesellschaft mit beschränkter Haftung
Example Music (Österreich) GmbH      Austria             GmbH                 Gesellschaft mit beschränkter Haftung
Example Music (CA) Inc.              Canada              Inc.                 Federal / provincial corporations
Example Music (US) LLC.              USA                 LLC.                 Standard US corporation suffix

Site  City / Location                Country             Subnet               Landline Range                      Mobile Range
----  ---------------------------    ---------           ------------------   ---------------------------------   -------------------
ABR   Aberdeen, Scotland             UK                  192.168.224.0/24     +44 1224 496 0xxx                   +44 7700 900 2xxx
BIR   Birmingham, England            UK                  192.168.121.0/24     +44 121  496 0xxx                   +44 7700 900 2xxx
BON   Bonn, West Germany (FRG)       DE                  192.168.228.0/24     +49 228  555  xxx                   +49 211  xxx xxxx
BRD   West Berlin (FRG)              DE                  192.168.113.0/24     +49 311  555  xxx                   +49 211  xxx xxxx
BRK   Brockville, Ontario            CA                  192.168.136.0/24     +1  613  555 6xxx                   +1  613  555 6xxx
CLY   Clydebank, Scotland            UK                  192.168.141.0/24     +44 141  496 00xx                   +44 770  090 5xxx
CPH   Copenhagen (København)         DK                  192.168.231.0/24     +45  00  000  xxx                   +45  2x  xxx xxx
DUN   Dundee, Scotland               UK                  192.168.138.0/24     +44 163  249 60xx                   +44 770  090 82xx
EDI   Edinburgh, Scotland            UK                  192.168.131.0/24     +44 131  496 0xxx                   +44 770  090 3xxx
FAX   Faxe, Danmark                  DK                  192.168.246.0/24     +45  00  000  xxx                   +45  2x  xxx xxx
GLA   Glasgow, Scotland              UK                  192.168.141.0/24     +44 141  496 01xx                   +44 770  009 4xxx
KGE   Køge, Danmark                  DK                  192.168.265.0/24     +45  00  000  xxx                   +45  2x  xxx xxx
KOR   Korsør, Danmark                DK                  192.168.238.0/24     +45  00  000  xxx                   +45  2x  xxx xxx
LIV   Liverpool, England             UK                  192.168.151.0/24     +44 151  496 0xxx                   +44 770  090 5xxx
LND   London, England                UK                  192.168.20.0/24      +44 207  496 0xxx / 01632 96x xxx   +44 770  090 0xxx
MCR   Manchester, England            UK                  192.168.161.0/24     +44 161  715 xxxx                   +44 770  090 6xxx
MIA   Miami, Florida                 US                  192.168.135.0/24     +1  305  555 xxxx                   +1  786  555 xxxx
MUN   Munich, West Germany           DE                  192.168.189.0/24     +49 893  555 33xx                   +49 893  555 99xx
NEW   Newcastle, England             UK                  192.168.191.0/24     +44 191  496 0xxx                   +44 770  090 9xxx
ODE   Odense, Danmark                DK                  192.168.126.0/24     +45  00  000  xxx                   +45  2x  xxx xxx
VIE   Vienna, Austria                AT                  192.168.xxx.0/24     +43 800  078  0xx                   +43 664  665 xxx

A note on phone numbers:

  - UK Phone Number Standards (Ofcom reserved ranges for fiction/testing):
    - Glasgow:                 0141 496 0xxx
    - Edinburgh:               0131 496 0xxx
    - London:                  0207 946 0xxx / 01632 96x xxx
    - Manchester:              0161 715 xxxx - Weatherfield (Yes, Coronation Street per the "Coronation Street test")
    - UK Wide Mobiles: Mobile: 0770 090 0xxx

  Per: https://en.wikipedia.org/wiki/Telephone_numbers_in_the_United_Kingdom?useskin=vector#Fictitious_numbers
  and: https://web.archive.org/web/20140216214400/http://stakeholders.ofcom.org.uk/telecoms/numbering/guidance-tele-no/numbers-for-drama

  - Denmark Testing Numbers:
    - Copenhagen landline: +45 0000 xxxx
    - Denmark mobile:      +45 2xxx xxxx

NB: I don't believe Denmark uses 0000 but this is not confirmed!

Device Role Codes:

BPS  = Badge Programming Station                              RAC  = Remote Access Controller (DRAC / iLO / BMC class)
CAM  = Security Camera                                        RDR  = Card Reader / Badge Reader
CLK  = Time Clock / Punch Clock                               RTR  = Router
COF  = Coffee Machine (Smart Appliance)                       SBC  = Session Border Controller
DON  = Donut Vending Machine                                  SRV  = Server
FWL  = Firewall Appliance                                     SUR  = Microsoft Surface Device
ILO  = Integrated Lights-Out (Server Management Controller)   SVR  = Server (Legacy / Alternate Code)
LAP  = Laptop (Windows)                                       SWI  = Network Switch
LCD  = LCD Wallboard / Information Display                    TAB  = Tablet
MAC  = macOS Desktop (iMac / Mac Mini)                        TEA  = Internet connected Coffee Pot / Tea machine RFC2324 compliant
MBP  = MacBook Pro                                            TVS  = Television / Digital Signage Display
MUS  = Music Workstation / Studio System                      VCU  = Video Conferencing Unit
NIX  = Unix/Linux/Solaris System                              WAP  = Wireless Access Point
PHN  = Mobile Phone                                           WKS  = Workstation (Desktop)
PRN  = Printer

Device Numbering: 000–999 (three digits, zero padded)

Examples:
  EXAWKSBON001  -> Bonn workstation
  EXASRVEDI003  -> Edinburgh server
  EXAWAPLND001  -> London Wi-Fi access point

Notes:
  - Not all devices are AD-aware
  - Some non-AD devices may still have service accounts
  - LAPS attributes exist only on supported Windows computers

If you need to bulk add another AD property ot demo data:

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
 Load-ADData or Load-DemoData
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

===========================================================================================
DSA-TUI Blåbær — Active Directory TUI Tool - Historical Build Notes and Change Log
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
 - Have other locations in Danmark along with different user properites e.g. Alan Wilder

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

1.8.3.0  (More cowbell)
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

1.8.6.0  (Colour my life)
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

2.3.7.1 (Bug fix)
  - Initialize-DirectoryEmoji based on date. Special days use different emojis, e.g. Dec 25th

2.3.7.2 (Bug fix and cleanup. more bands)
  - Fix some erroneous $Globals that were left in
  - Fix group membership modal
  - Load the terminal icons module, if found
  - Advise user about the Nerd fonts for directory icons in pwsh 7+
  - Add other devices besides comptuer and printers. While not strictly part of AD corporate networks owuld have this.
  - Add Altered Images and Eurythmics Bands
  - Add UPNs and SamAccountName to align with real world AD
  - Bring back thw search/lookup tab for users... Spis lige brød til!

2.3.8.0 (Code reflow)
  - If terminal icons and a suitable nerd font is avilable, it will be used for prefixes, rather than (U), (G), etc.
  - REMOVED Get-CleanObjectInfo function which is no longer needed by switching to .Tag to detect type
  - Simplify calls to Objects and add icons for them too
  - Rework startup of app to avoid "hanging" by repainting status bar as enumeration progresses

2.3.8.21  (Bug fix and code consolidaiton)
  - Fix Theme selector and bug in main window background arising from it
  - Upgrade Show-Modal to also handle Yes/No returning 0 for YES and 1 for NO (The order the buttons show)
  - Even more themes!
  - Create a New-PropertiesDialog to redcue repetitve code. Part of a wider refsctoring effort for code maintainability
  - With the above function rework, rewrtie Show-{Computer|DC|OU|User}-Properties to use it
  - Show-EditGroupMembershipDialog to reduce code reuse even more
  - Add in some "fun" things, such as xpired LAPS and Laptops that haven't been "on" for months, etc.
  - Clean up and merge a number of funcitons to reduce code re-use
  - https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/cc732101(v=ws.11)
    new tools, also allows import (or export) their own data. This code is use as is, it does not syntax check the CSV!

3.0.0.29 (Cleanup and reflow many areas of code)
  - Massive consolidation / feature expansion session. Reduced code duplication with unified functions, added bulk
    operations and auditing features.
  - Manage-FilterStatusLabel
    - Unified Create-FilterStatusLabel + Update-FilterStatusLabel
    - Single function with -Action 'Create'/'Update' parameter
    - Supports both standalone and in-panel creation with -InPanel switch
    - Replaced: Create-FilterStatusLabel, Update-FilterStatusLabel
  - Manage-Spinner
   - Unified Start-Spinner + Stop-Spinner
   - Single function with -Action 'Start'/'Stop' parameter
   - Message parameter validation when starting
   - Replaced: Start-Spinner, Stop-Spinner
  - Manage-Selecti
   - Unified Select-AllObjects + Deselect-AllObjects
   - Single function with -Action 'SelectAll'/'DeselectAll' parameter
   - Replaced: Select-AllObjects, Deselect-AllObjects
  - Get-Theme
   - Combined Get-Theme + Dump-ColourScheme into single function
   - Added -Dump switch for debugging theme colours
   - Can dump current theme or specific theme: Get-Theme -Mode "matrix" -Dump
   - All 19 themes preserved (British, Class91, Dark, DB-1980s, etc.)
   - Replaced: Separate Dump-ColourScheme function
  - Apply-ObjectChanges
   - Unified Apply-UserChanges, Apply-GroupChanges, Apply-OUChanges into Single function handles User/Group/OU/Computer
     with -ObjectType parameter
   - Handles both DemoMode and Production with all field types
   - User fields: 13 attributes including DisplayName, Email, Phone, Address
   - Group fields: Description, Email, ManagedBy
   - OU fields: Name (with rename + user reference updates), Description
   - Computer fields: Description, Location
   - Replaced: Apply-UserChanges, Apply-GroupChanges, Apply-OUChanges, Update-UserObjectFromFields
  - Invoke-ObjectOperation
   - Unified Show-DeleteObjectDialog, Show-MoveObjectDialog, Invoke-BulkMove
   - Single function with -Operation Move/Delete and -IsBulk switch
   - DELETE: Auto-detects object type, double confirmation, removes from arrays/AD
   - MOVE: Validates moveability, shows OU picker, bulk/single UI modes
   - Supports Users, Groups, Computers, OUs, Domain Controllers
   - Success/failure tracking with detailed error reporting
   - Replaced: Show-DeleteObjectDialog, Show-MoveObjectDialog, Invoke-BulkMove
  - Extracted AD Health Functions (582 lines total - MAJOR REFACTOR)
   - Broke 400+ line monolith Get-ADHealth into modular components
   - Created 7 standalone test functions + unified orchestrator
     - Invoke-ExternalCommand (helper for safe command execution)
     - Test-DCStatus (Domain Controller health checks)
     - Test-ADReplication (replication status with repadmin/Get-ADReplicationFailure)
     - Test-ADDnsRecords (DNS SRV records and service status)
     - Test-SysvolHealth (SYSVOL shares and DFSR status)
     - Test-FSMORoles (FSMO role holders and reachability)
     - Test-GPOHealth (GPO enumeration and version checks)
     - Check-ADHealth (main orchestrator with tabbed UI)
     - Each check independently callable
     - Reusable in other scripts
     - Easier testing and maintenance
     - Renamed from Get-ADHealth to Check-ADHealth for clarity
  - Show-ADToolsModal
    - Added: csvde.exe, nslookup.exe, ping.exe, net.exe, tracert.exe
    - Shows tool availability with descriptions and Microsoft docs links
    - Note about future CSV import feature for csvde
  - Manage-AccountStatus
    - Bulk Enable/Disable for users and computers
    - Validates object types before processing
    - Optional reason logging
    - Demo + Production mode support
    - Success/failure tracking with error summary
    - Usage: Manage-AccountStatus -Objects $users -Action 'Disable' -Reason "Offboarding"
  - Set-BulkAttribute
    - Change any attribute across multiple objects
    - Interactive dialog mode with attribute picker
    - Supports Users (13 attributes), Groups (3), Computers (2)
    - ChangePasswordAtLogon - Force password change at next login
    - ResetPassword - Bulk password reset with random generation
    - Random password generator (12 chars: A-Z, a-z, 0-9, special)
    - Context-sensitive help in dialog
    - Demo mode logs generated passwords
      ATTRIBUTES SUPPORTED: Users: DisplayName, Description, EmailAddress, Title, Department, Company, Manager,
                            OfficePhone, MobilePhone, StreetAddress, City, PostalCode, Country, ChangePasswordAtLogon,
                            ResetPassword
                            Groups: Description, Email, ManagedBy
                            Computers: Description, Location
  - Find-StaleAccounts
    - Find accounts not used in X days (default: 90)
    - Interactive parameter selection dialog with radio buttons
    - Supports Users, Computers, or Both
    - Shows last logon date and days inactive
    - Export to CSV functionality
    - CAN DISABLE DIRECTLY from results dialog
    - Demo mode: pseudo-random last logon dates based on name hash
    - Production: queries AD LastLogonTimeStamp
    - Usage: Find-StaleAccounts -ShowDialog (shows parameter picker first)
  - Copy-ADObject
    - Template-based user/group cloning
    - Interactive dialog with auto-generation
    - Auto-generates SamAccountName from display name (firstname.lastname)
    - Optional group membership/member copying
    - Copies all relevant fields (Department, Title, Company, Manager, etc.)
    - Demo + Production mode support
    - Usage: Copy-ADObject -SourceObject $templateUser -ShowDialog
  - Manage-DemoData
    - Import/Export demo data to CSV
    - DEMO MODE ONLY - blocks production mode with warning
    - Uses Show-FileBrowserDialog for file selection
    - Exports Users, Groups, Computers with all fields
    - Auto-rebuilds tree after import
    - UTF-8 encoding for international characters
    - Compatible with Microsoft csvde.exe format
    - Usage: Manage-DemoData -Action 'Export'/'Import'
  - New-PropertiesDialog
    - Tab builder pattern for property dialogs
    - Eliminates boilerplate in User/Group/OU/Computer/DC properties
    - Centralized Apply logic with shared state hashtables
    - Fixed Debug-Log scope issues (uses plain calls, not Get-Command)

3.0.0.31  (Bug fixes)
  - Users are now enumerated proprly and shown in real AD.
  - LAPS Password modal now shows correctly
  - Group OU, Computer and DC properties all now have a properties tab too
  - Add the AD tools into the AD Health modal. This also checks AD health
  - Rework AD Health to be more feature rich
  - Change domain worked first time to type in, didn't refresh anything, then crashed is now fixed
  - Change domain now falls back ot previous domain on failed to load a new domain

-------------------------------------------------------------------------------
TODO / COME BACK TO
-------------------------------------------------------------------------------

REMAINING FEATURES TO IMPLEMENT:

  - Account Expiration Management
  - Group Membership Report/Comparison
  - OU Statistics/Summary
  - Audit Log Viewer (demo mode fake logs)
  - Misisng AD module is non fatal BUT if it's not installed, a global needs to not let users do stupid stuff
  - show locked users actually shows computers under maintenence, which is... nice but not what you're asking for - add both
  - Did the right click popup go away or is it broken...? Yes. Fix it later
  - Groups aren't enumerated properly in produciton. Likely a logic bug in detetion
  - Reset-LapsPassword only works locally, but I know you can do remote LAPS rotation. Find out how add ot LAPS window
  - Menu entries for Set-BulkAttribute and Find-StaleAccounts
    (Need to use existing helper functions for selection conversion)
  - AD Health tabs like Group Policy and domain controllers the Tab tab pane could be smaller with a search box in them to help out
  - dns queries ought ot unr automatically

 BUGS:
  - Groups are not enumerated propely in the tree. Much like the user bug, it's likely properties. Computers is also likely affected

===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain,  ## User can specify domain
  [ValidateSet("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")]
  [string]$Theme
)

## Define the build version, project and code names once only - up here to ease patching. The rest at main
$Script:ProjectName  = "DSA-TUI pwsh dsa.msc TUI"
$Script:FruitName    = "Blåbær"
$Script:BuildVersion = "3.0.0.29"

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
    return $false
    Debug-Log "If yuo would like terminal icons, be sure to install this: https://github.com/jpawlowski/nerd-fonts-installer-PS" -Type "Info"
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

function Build-MainMenu {
  [CmdletBinding()]
  param()

  Debug-Log ": Building main menu..." -Type "Info"

  ## -------------------------{ Menu Items -------------------------
  $mFile         = [Terminal.Gui.MenuItem]::new("_Exit","Exit application (F10)",[Action]{ [Terminal.Gui.Application]::RequestStop() })
  $mNew          = [Terminal.Gui.MenuItem]::new("New Object","Create a new object (F3)",[Action]{ Show-NewObjectWizard })
  $mProps        = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })

  ## In Build-MainMenu function:
  $mDemoExport   = [Terminal.Gui.MenuItem]::new("_Export Demo Data", "Export demo data to CSV", [Action]{ Manage-DemoData -Action 'Export' })
  $mDemoImport   = [Terminal.Gui.MenuItem]::new("_Import Demo Data", "Import demo data from CSV", [Action]{ Manage-DemoData -Action 'Import' })

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

  $mQuickFilter       = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{ Show-QuickFilterDialog })
  $mSelectionMode     = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{ Toggle-SelectionMode })
  $mSelectAll         = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{ Manage-Selection -Action 'SelectAll' })
  $mDeselectAll       = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{ Manage-Selection -Action 'DeselectAll' })
  $mBulkDisable       = [Terminal.Gui.MenuItem]::new("_Disable Selected", "Disable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Disable' })
  $mBulkEnable        = [Terminal.Gui.MenuItem]::new("_Enable Selected", "Enable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Enable' })
  $mBulkAddGroup      = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{ Invoke-BulkAddToGroup })
  $mBulkEdit          = [Terminal.Gui.MenuItem]::new("_Edit Attribute (Bulk)", "Change one attribute across selected objects", [Action]{$objects = Get-SelectedObjectsAsObjects ; Set-BulkAttribute -Objects $objects -ShowDialog})
  $mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator (F2)",[Action]{ Generate-RandomPassword })
  $mLAPSPasswords     = [Terminal.Gui.MenuItem]::new("_LAPS Passwords","Lookup LAPS Creds (Fx)",[Action]{ Show-LAPSSearchModal })
  $mADHealth          = [Terminal.Gui.MenuItem]::new("_AD Health Status", "AD Health And Replication Status", [Action]{ Show-ADHealthDialog })
  $mShortcuts         = [Terminal.Gui.MenuItem]::new("_Shortcuts","Keyboard shortcuts (F1)",[Action]{ Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF10 - Quit" })
  $mAboutDSATUI       = [Terminal.Gui.MenuItem]::new("_About","About $($Script:ProjectName)",[Action]{ Show-Modal "About" "$($Script:ProjectName)`n`nCodename: $($Script:FruitName)`nv$($Script:BuildVersion) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600)" })
  $mWhyBlaabaer       = [Terminal.Gui.MenuItem]::new("Why _Blaabaer?","Why the $($Script:FruitName) codename?",[Action]{ Show-BlaabaerInfo })
  $mTheme             = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme (F6)",[Action]{ Show-ThemeSelector })

  $mCopyTemplate      = [Terminal.Gui.MenuItem]::new("Copy as _Template", "Create new object based on selected object", [Action]{
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
  $menu = [Terminal.Gui.MenuBar]::new(@(
    [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mDemoExport, $mDemoImport, $mTheme, $mFile)),

    [Terminal.Gui.MenuBarItem]::new("_Action", @($mNew, $mProps, $mQuickFilter, $mUndo, $mChangeDomain, $mLAPSPasswords, $mPasswordGenerator, $mChangeDC, $mSearchAD, $mADHealth, $mCopyTemplate)),
    [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode, $mSelectAll, $mDeselectAll, $mBulkEdit, $mBulkAddGroup, $mBulkEnable, $mBulkDisable)),
    [Terminal.Gui.MenuBarItem]::new("_About", @($mShortcuts, $mAboutDSATUI, $mWhyBlaabaer))
  ))

  Debug-Log ": Main menu created successfully" -Type "Info"
  return $menu
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
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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
        Name = $user.Name
        SamAccountName = $user.SamAccountName
        Enabled = $user.Enabled
        Disabled = -not $user.Enabled
        LastLogon = $lastLogon
        DaysSinceLogon = $daysAgo
        OU = $user.DistinguishedName -replace '^CN=[^,]+,', ''
        Department = $user.Department
        Title = $user.Title
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
        $daysAgo = ($hash % 365) + 1  # 1-365 days ago
        $lastLogon = (Get-Date).AddDays(-$daysAgo)

        if ($lastLogon -lt $cutoffDate) {
          $staleAccounts += [PSCustomObject]@{
          ObjectType = 'Computer'
          Name = $computer.Name
          SamAccountName = $computer.SamAccountName
          Enabled = $computer.Enabled
          Disabled = $computer.Disabled
          LastLogon = $lastLogon
          DaysSinceLogon = $daysAgo
          OU = 'Computers'
          OperatingSystem = $computer.OperatingSystem
          Location = $computer.Location
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
        ObjectType = 'Computer'
        Name = $computer.Name
        SamAccountName = $computer.SamAccountName
        Enabled = $computer.Enabled
        Disabled = -not $computer.Enabled
        LastLogon = $lastLogon
        DaysSinceLogon = $daysAgo
        OU = $computer.DistinguishedName -replace '^CN=[^,]+,', ''
        OperatingSystem = $computer.OperatingSystem
        Location = $computer.Location
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

        if ($displayItems.Count -eq 0) {
            $displayItems = @("(No stale accounts found)")
        }

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
            if ($selected.ObjectType -eq 'User') {
                $obj = $Script:Users | Where-Object { $_.SamAccountName -eq $selected.SamAccountName } | Select-Object -First 1
            } else {
                $obj = $Script:Computers | Where-Object { $_.SamAccountName -eq $selected.SamAccountName } | Select-Object -First 1
            }

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
        $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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
    $isUser = $SourceObject.PSObject.Properties.Match('SamAccountName') -and
              -not $SourceObject.PSObject.Properties.Match('ComputerType')
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

        if ($isUser) {
            $fieldsList = "Department, Title, Company, Manager, OU, Phone, Address"
        } else {
            $fieldsList = "Description, ManagedBy, OU"
        }
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
                    NewName = $name
                    SamAccountName = $sam
                    EmailAddress = $email
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
        $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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
    $emailAddress = $PSBoundParameters['EmailAddress']

    ## Create new object based on template
    try {
        if ($Script:DemoMode) {
            ## ==================== DEMO MODE - COPY USER ====================
            if ($isUser) {
                ## Generate SamAccountName if not provided
                if (-not $samAccountName) {
                    if ($NewName -match '^(\w+)\s+(\w+)') {
                        $samAccountName = "$($Matches[1]).$($Matches[2])".ToLower()
                    } else {
                        $samAccountName = $NewName.ToLower().Replace(' ', '.')
                    }
                }

                ## Generate email if not provided
                if (-not $emailAddress -and $SourceObject.EmailAddress) {
                    $domain = $SourceObject.EmailAddress -replace '^[^@]+@', ''
                    $emailAddress = "${samAccountName}@${domain}"
                }

                ## Create new user object
                $newUser = [PSCustomObject]@{
                    Name = $NewName
                    SamAccountName = $samAccountName
                    DisplayName = $NewName
                    EmailAddress = $emailAddress
                    Enabled = $true
                    Disabled = $false
                    LockedOut = $false
                    OU = $SourceObject.OU
                    Groups = if ($CopyMemberships) { $SourceObject.Groups } else { @() }
                    MemberOf = if ($CopyMemberships) { $SourceObject.MemberOf } else { @() }
                    Title = $SourceObject.Title
                    Department = $SourceObject.Department
                    Company = $SourceObject.Company
                    Manager = $SourceObject.Manager
                    OfficePhone = $SourceObject.OfficePhone
                    MobilePhone = $SourceObject.MobilePhone
                    StreetAddress = $SourceObject.StreetAddress
                    City = $SourceObject.City
                    PostalCode = $SourceObject.PostalCode
                    Country = $SourceObject.Country
                    UserPrincipalName = "${samAccountName}@$($Script:CurrentDomain)"
                    DistinguishedName = "CN=$NewName,$($SourceObject.DistinguishedName -replace '^CN=[^,]+,', '')"
                    Domain = $Script:CurrentDomain
                }

                $Script:Users += $newUser
                $Script:rawUsers += $newUser

                Debug-Log ": Created user '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
                if ($CopyMemberships) {
                    Debug-Log ":   Copied $($newUser.Groups.Count) group memberships" -Type "Info"
                }

                Show-Modal "User Created" "Successfully created user '$NewName'`n`nLogin: $samAccountName`nEmail: $emailAddress$(if ($CopyMemberships) { "`n`nCopied $($newUser.Groups.Count) group memberships" } else { '' })"

            ## ==================== DEMO MODE - COPY GROUP ====================
            } else {
                ## Create new group object
                $newGroup = [PSCustomObject]@{
                    Name = $NewName
                    Description = $SourceObject.Description
                    Email = $SourceObject.Email
                    ManagedBy = $SourceObject.ManagedBy
                    Members = if ($CopyMemberships) { $SourceObject.Members } else { @() }
                    MemberOf = $SourceObject.MemberOf
                    DistinguishedName = "CN=$NewName,$($SourceObject.DistinguishedName -replace '^CN=[^,]+,', '')"
                    Domain = $Script:CurrentDomain
                }

                $Script:Groups += $newGroup
                $Script:rawDemoGroups += $newGroup

                Debug-Log ": Created group '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
                if ($CopyMemberships) {
                    Debug-Log ":   Copied $($newGroup.Members.Count) members" -Type "Info"
                }

                Show-Modal "Group Created" "Successfully created group '$NewName'$(if ($CopyMemberships) { "`n`nCopied $($newGroup.Members.Count) members" } else { '' })"
            }

        } else {
            ## ==================== PRODUCTION MODE ====================
            if ($isUser) {
                ## Generate SamAccountName if not provided
                if (-not $samAccountName) {
                    if ($NewName -match '^(\w+)\s+(\w+)') {
                        $samAccountName = "$($Matches[1]).$($Matches[2])".ToLower()
                    } else {
                        $samAccountName = $NewName.ToLower().Replace(' ', '.')
                    }
                }

                ## Create user params
                $params = @{
                    Name = $NewName
                    SamAccountName = $samAccountName
                    DisplayName = $NewName
                    Path = $SourceObject.DistinguishedName -replace '^CN=[^,]+,', ''
                    Enabled = $true
                    ErrorAction = 'Stop'
                }

                ## Copy standard fields
                if ($SourceObject.Title) { $params['Title'] = $SourceObject.Title }
                if ($SourceObject.Department) { $params['Department'] = $SourceObject.Department }
                if ($SourceObject.Company) { $params['Company'] = $SourceObject.Company }
                if ($SourceObject.Manager) { $params['Manager'] = $SourceObject.Manager }
                if ($SourceObject.OfficePhone) { $params['OfficePhone'] = $SourceObject.OfficePhone }
                if ($SourceObject.StreetAddress) { $params['StreetAddress'] = $SourceObject.StreetAddress }
                if ($SourceObject.City) { $params['City'] = $SourceObject.City }
                if ($SourceObject.PostalCode) { $params['PostalCode'] = $SourceObject.PostalCode }
                if ($SourceObject.Country) { $params['Country'] = $SourceObject.Country }
                if ($emailAddress) { $params['EmailAddress'] = $emailAddress }

                ## Create user
                New-ADUser @params

                ## Copy group memberships
                if ($CopyMemberships) {
                    $sourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceObject.SamAccountName | Where-Object { $_.Name -ne 'Domain Users' }
                    foreach ($group in $sourceGroups) {
                        Add-ADGroupMember -Identity $group.SamAccountName -Members $samAccountName -ErrorAction SilentlyContinue
                    }
                    Debug-Log ":   Copied $($sourceGroups.Count) group memberships" -Type "Info"
                }

                Debug-Log ": Created user '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
                Show-Modal "User Created" "Successfully created user '$NewName' in AD$(if ($CopyMemberships) { "`n`nCopied $($sourceGroups.Count) group memberships" } else { '' })"

            } else {
                ## Create group
                $params = @{
                    Name = $NewName
                    Path = $SourceObject.DistinguishedName -replace '^CN=[^,]+,', ''
                    GroupScope = 'Global'
                    ErrorAction = 'Stop'
                }

                if ($SourceObject.Description) { $params['Description'] = $SourceObject.Description }
                if ($SourceObject.ManagedBy) { $params['ManagedBy'] = $SourceObject.ManagedBy }

                New-ADGroup @params

                ## Copy members
                if ($CopyMemberships) {
                    $sourceMembers = Get-ADGroupMember -Identity $SourceObject.Name
                    foreach ($member in $sourceMembers) {
                        Add-ADGroupMember -Identity $NewName -Members $member.SamAccountName -ErrorAction SilentlyContinue
                    }
                    Debug-Log ":   Copied $($sourceMembers.Count) members" -Type "Info"
                }

                Debug-Log ": Created group '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
                Show-Modal "Group Created" "Successfully created group '$NewName' in AD$(if ($CopyMemberships) { "`n`nCopied $($sourceMembers.Count) members" } else { '' })"
            }
        }

        ## Refresh UI
        Refresh-Data -domain $Script:CurrentDomain
        Build-Tree -domain $Script:CurrentDomain
        if ($Script:FilterStatusLabel) {
            Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
        }

    } catch {
        Debug-Log ": Failed to create $objectType '$NewName': $($_.Exception.Message)" -Type "Error"
        Show-Modal "Creation Failed" "Failed to create $objectType '$NewName':`n`n$($_.Exception.Message)"
    }
}

## ==================== DEMO DATA IMPORT/EXPORT ====================
function Manage-DemoData {
    <#
    .SYNOPSIS
    Import or export demo data to/from CSV files with analysis

    .PARAMETER Action
    'Import' to load CSV data, 'Export' to save current demo data

    .PARAMETER FilePath
    Optional - specify file path directly. If not provided, shows file browser

    .EXAMPLE
    Manage-DemoData -Action 'Export'

    .EXAMPLE
    Manage-DemoData -Action 'Import' -FilePath "C:\data\ad-export.csv"

    .NOTES
    WARNING: This function works ONLY in demo mode.
    CSV format: Active Directory CSV export (csvde.exe compatible)
    Analyzes and counts all object types before import.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Import', 'Export')]
        [string]$Action,

        [string]$FilePath
    )

    ## Safety check - demo mode only
    if (-not $Script:DemoMode) {
        Show-Modal "Production Mode" "Demo data import/export is ONLY available in demo mode.`n`nCurrent mode: Production`n`nRestart in demo mode to use this feature."
        Debug-Log ": Attempted demo data $Action in production mode - blocked" -Type "Warn"
        return
    }

    Debug-Log ": Demo data $Action initiated" -Type "Info"

    ## ==================== EXPORT ====================
    if ($Action -eq 'Export') {
        ## If no path provided, show save dialog
        if (-not $FilePath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

            ## Show file browser for save location
            $FilePath = Show-FileBrowserDialog -StartDir "." -Title "Export Demo Data - Select Location" -Filter @("*.csv", "*.txt")

            if (-not $FilePath) {
                Debug-Log ": Export cancelled by user" -Type "Info"
                return
            }

            ## Ensure .csv extension
            if ($FilePath -notmatch '\.(csv|txt)$') {
                $FilePath = "$FilePath-$timestamp.csv"
            }
        }

        try {
            ## Combine all demo data into exportable format
            $exportData = @()

            ## Export Users
            foreach ($user in $Script:Users) {
                $exportData += [PSCustomObject]@{
                    ObjectType = 'User'
                    Name = $user.Name
                    SamAccountName = $user.SamAccountName
                    DisplayName = $user.DisplayName
                    EmailAddress = $user.EmailAddress
                    Enabled = $user.Enabled
                    Disabled = $user.Disabled
                    LockedOut = $user.LockedOut
                    OU = ($user.OU -join ';')
                    Groups = ($user.Groups -join ';')
                    MemberOf = ($user.MemberOf -join ';')
                    Title = $user.Title
                    Department = $user.Department
                    Company = $user.Company
                    Manager = $user.Manager
                    OfficePhone = $user.OfficePhone
                    MobilePhone = $user.MobilePhone
                    StreetAddress = $user.StreetAddress
                    City = $user.City
                    PostalCode = $user.PostalCode
                    Country = $user.Country
                    UserPrincipalName = $user.UserPrincipalName
                    DistinguishedName = $user.DistinguishedName
                    Domain = $user.Domain
                }
            }

            ## Export Groups
            foreach ($group in $Script:Groups) {
                $exportData += [PSCustomObject]@{
                    ObjectType = 'Group'
                    Name = $group.Name
                    Description = $group.Description
                    Email = $group.Email
                    ManagedBy = $group.ManagedBy
                    Members = ($group.Members -join ';')
                    MemberOf = ($group.MemberOf -join ';')
                    DistinguishedName = $group.DistinguishedName
                    Domain = $group.Domain
                }
            }

            ## Export Computers
            foreach ($computer in $Script:Computers) {
                $exportData += [PSCustomObject]@{
                    ObjectType = 'Computer'
                    Name = $computer.Name
                    SamAccountName = $computer.SamAccountName
                    Enabled = $computer.Enabled
                    Disabled = $computer.Disabled
                    OperatingSystem = $computer.OperatingSystem
                    ComputerType = $computer.ComputerType
                    Location = $computer.Location
                    DistinguishedName = $computer.DistinguishedName
                    Domain = $computer.Domain
                    Description = $computer.Description
                }
            }

            ## Write to CSV
            $exportData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8

            $count = $exportData.Count
            Debug-Log ": Exported $count objects to $FilePath" -Type "Success"
            Show-Modal "Export Complete" "Successfully exported $count objects to:`n`n$FilePath"

        } catch {
            Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
            Show-Modal "Export Failed" "Failed to export demo data:`n`n$($_.Exception.Message)"
        }
    }

    ## ==================== IMPORT ====================
    if ($Action -eq 'Import') {
        ## If no path provided, show file browser
        if (-not $FilePath) {
            $FilePath = Show-FileBrowserDialog -StartDir "." -Title "Import AD CSV Data - Select File" -Filter @("*.csv", "*.txt")

            if (-not $FilePath) {
                Debug-Log ": Import cancelled by user" -Type "Info"
                return
            }
        }

        ## Verify file exists
        if (-not (Test-Path $FilePath)) {
            Show-Modal "File Not Found" "Cannot find file:`n`n$FilePath"
            Debug-Log ": Import failed - file not found: $FilePath" -Type "Error"
            return
        }

        try {
            ## Import and analyze CSV first
            Debug-Log ": Analyzing CSV file: $FilePath" -Type "Info"
            $importData = Import-Csv -Path $FilePath -Encoding UTF8

            if (-not $importData -or $importData.Count -eq 0) {
                throw "CSV file is empty or invalid"
            }

            ## Count objects by type (using objectClass from AD CSV)
            $stats = @{
                Users = 0
                Groups = 0
                Computers = 0
                DomainControllers = 0
                Servers = 0
                OrganizationalUnits = 0
                Contacts = 0
                PrintQueues = 0
                Other = 0
                Total = $importData.Count
            }

            foreach ($row in $importData) {
                $objectClass = $row.objectClass

                switch -Regex ($objectClass) {
                    '^user$' {
                        $stats.Users++
                    }
                    '^group$' {
                        $stats.Groups++
                    }
                    '^computer$' {
                        ## Distinguish DCs from regular computers/servers
                        if ($row.userAccountControl -match '532480|8192' -or
                            $row.servicePrincipalName -match 'E3514235-4B06-11D1-AB04-00C04FC2DCD2') {
                            $stats.DomainControllers++
                        } elseif ($row.operatingSystem -match 'Server') {
                            $stats.Servers++
                        } else {
                            $stats.Computers++
                        }
                    }
                    '^organizationalUnit$' {
                        $stats.OrganizationalUnits++
                    }
                    '^contact$' {
                        $stats.Contacts++
                    }
                    '^printQueue$' {
                        $stats.PrintQueues++
                    }
                    '^(domainDNS|container|builtinDomain)$' {
                        ## Structural objects - skip silently
                    }
                    default {
                        $stats.Other++
                        Debug-Log ": Unknown object class: $objectClass ($($row.name))" -Type "Warn"
                    }
                }
            }

            Debug-Log ": CSV Analysis - Users: $($stats.Users), Groups: $($stats.Groups), Computers: $($stats.Computers), DCs: $($stats.DomainControllers), OUs: $($stats.OrganizationalUnits)" -Type "Info"

            ## Build summary message with nice formatting
            $summary = @"
CSV File Analysis
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

File: $(Split-Path $FilePath -Leaf)
Total Objects: $($stats.Total)

Object Breakdown:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Users:                 $($stats.Users)
Groups:                $($stats.Groups)
Computers:             $($stats.Computers)
Domain Controllers:    $($stats.DomainControllers)
Servers:               $($stats.Servers)
Organizational Units:  $($stats.OrganizationalUnits)
Contacts:              $($stats.Contacts)
Print Queues:          $($stats.PrintQueues)
Other Objects:         $($stats.Other)

⚠️  WARNING: This will REPLACE all current demo data.

Continue with import?
"@

            ## Show analysis and confirm
            $confirm = Show-Modal "CSV Import Analysis" $summary -YesNo

            if ($confirm -ne 0) {
                Debug-Log ": Import cancelled by user after analysis" -Type "Info"
                return
            }

            ## Clear existing demo data
            $Script:rawDCs = @()
            $Script:Users = @()
            $Script:rawUsers = @()
            $Script:Groups = @()
            $Script:rawDemoGroups = @()
            $Script:Computers = @()
            $Script:rawComputers = @()

            ## Import objects by type
            $userCount = 0
            $groupCount = 0
            $computerCount = 0
            $dcCount = 0
            $ouCount = 0

            Debug-Log ": Beginning CSV import..." -Type "Info"

            ## ==================== INTERNAL HELPER FUNCTIONS ====================

            function ParseOUPath {
                param([string]$DN)

                if ([string]::IsNullOrWhiteSpace($DN)) { return @() }

                $ouParts = @()
                $DN -split '(?<!\\),' | Where-Object { $_ -match '^OU=' } | ForEach-Object {
                    if ($_ -match '^OU=(.+)$') { $ouParts += $matches[1] }
                }
                [array]::Reverse($ouParts)
                return $ouParts
            }

            function ExtractDomain {
                param([string]$DN)

                if ([string]::IsNullOrWhiteSpace($DN)) { return 'example.com' }

                $dcParts = $DN -split '(?<!\\),' | Where-Object { $_ -match '^DC=' } | ForEach-Object {
                    if ($_ -match '^DC=(.+)$') { $matches[1] }
                }

                if ($dcParts) {
                    return ($dcParts -join '.')
                } else {
                    return 'example.com'
                }
            }

            foreach ($row in $importData) {
                $objectClass = $row.objectClass

                switch -Regex ($objectClass) {
                    '^user$' {
                        ## Parse user from CSV
                        $dn = $row.distinguishedName
                        $ouPath = ParseOUPath -DN $dn
                        $domain = ExtractDomain -DN $dn

                        $user = [PSCustomObject]@{
                            Name = $row.name
                            SamAccountName = $row.sAMAccountName
                            DisplayName = $row.displayName
                            EmailAddress = $row.mail
                            UserPrincipalName = $row.userPrincipalName
                            Enabled = -not ([int]$row.userAccountControl -band 0x0002)
                            Disabled = ([int]$row.userAccountControl -band 0x0002) -ne 0
                            LockedOut = ([int]$row.userAccountControl -band 0x0010) -ne 0
                            OU = $ouPath
                            Groups = if ($row.memberOf) { ($row.memberOf -split ';' | ForEach-Object { if ($_ -match 'CN=([^,]+)') { $matches[1] } }) } else { @() }
                            MemberOf = if ($row.memberOf) { ($row.memberOf -split ';') } else { @() }
                            Title = $row.title
                            Department = $row.description
                            Company = ''
                            Manager = if ($row.manager -match 'CN=([^,]+)') { $matches[1] } else { '' }
                            OfficePhone = $row.telephoneNumber
                            MobilePhone = $row.mobile
                            StreetAddress = ''
                            City = ''
                            PostalCode = ''
                            Country = $row.countryCode
                            DistinguishedName = $dn
                            Domain = $domain
                        }
                        $Script:Users += $user
                        $Script:rawUsers += $user
                        $userCount++
                    }

                    '^group$' {
                        $group = [PSCustomObject]@{
                            Name = $row.name
                            Description = $row.description
                            Email = $row.mail
                            ManagedBy = if ($row.managedBy -match 'CN=([^,]+)') { $matches[1] } else { '' }
                            Members = if ($row.member) { ($row.member -split ';') } else { @() }
                            MemberOf = if ($row.memberOf) { ($row.memberOf -split ';') } else { @() }
                            DistinguishedName = $row.distinguishedName
                            Domain = 'example.com'
                        }
                        $Script:Groups += $group
                        $Script:rawDemoGroups += $group
                        $groupCount++
                    }

                    '^computer$' {
                        ## Check if DC
                        if ($row.userAccountControl -match '532480|8192' -or
                            $row.servicePrincipalName -match 'E3514235-4B06-11D1-AB04-00C04FC2DCD2') {
                            ## It's a DC
                            $dc = @{
                                Name = $row.name
                                SamAccountName = $row.sAMAccountName
                                DNSHostName = $row.dNSHostName
                                Site = 'Default-First-Site-Name'
                                Location = ''
                                Domain = 'example.com'
                                Forest = 'example.com'
                                OS = $row.operatingSystem
                                OperatingSystemVersion = $row.operatingSystemVersion
                                IPv4Address = ''
                                Enabled = $true
                                IsGlobalCatalog = $row.servicePrincipalName -match 'GC/'
                                FSMORoles = @()
                                LastReplication = Get-Date
                                ReplicationHealth = 'Healthy'
                                LastBoot = Get-Date
                            }
                            $Script:rawDCs += $dc
                            $dcCount++
                        } else {
                            ## Regular computer
                            $computer = [PSCustomObject]@{
                                Name = $row.name
                                SamAccountName = $row.sAMAccountName
                                Enabled = -not ([int]$row.userAccountControl -band 0x0002)
                                Disabled = ([int]$row.userAccountControl -band 0x0002) -ne 0
                                OperatingSystem = $row.operatingSystem
                                OperatingSystemVersion = $row.operatingSystemVersion
                                ComputerType = if ($row.operatingSystem -match 'Server') { 'Server' } else { 'Workstation' }
                                Location = ''
                                Description = $row.description
                                DistinguishedName = $row.distinguishedName
                                Domain = 'example.com'
                            }
                            $Script:Computers += $computer
                            $Script:rawComputers += $computer
                            $computerCount++
                        }
                    }

                    '^organizationalUnit$' {
                        $ouCount++
                        ## OUs are structural, don't need explicit import
                    }

                    '^(domainDNS|container|builtinDomain)$' {
                        ## Structural/system objects - skip silently
                    }
                }
            }

            ## Update totals
            $Script:ADObjects = $Script:Users + $Script:Groups + $Script:Computers + $Script:rawDCs
            $Script:DCs = $Script:rawDCs

            Debug-Log ": Import complete - Users: $userCount, Groups: $groupCount, Computers: $computerCount, DCs: $dcCount, OUs: $ouCount" -Type "Success"

            ## Show completion summary
            $completionMsg = @"
Import Complete!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Imported Objects:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Users:                $userCount
Groups:               $groupCount
Computers:            $computerCount
Domain Controllers:   $dcCount
Organizational Units: $ouCount

Total Objects: $($Script:ADObjects.Count)

The tree will be rebuilt automatically.
"@

            Show-Modal "Import Complete" $completionMsg

            ## Rebuild tree with new data
            Debug-Log ": Rebuilding tree with imported data..." -Type "Info"
            Refresh-Data -domain $Script:CurrentDomain
            $rootNode = Build-Tree -domain $Script:CurrentDomain
            if ($rootNode) {
                $Script:tree.ClearObjects()
                $Script:tree.AddObject($rootNode)
            }

        } catch {
            Debug-Log ": Import failed: $($_.Exception.Message)" -Type "Error"
            Show-Modal "Import Failed" "Failed to import demo data:`n`n$($_.Exception.Message)`n`nNote: CSV must be from AD CSV export tool (csvde.exe).`nIf you mess it up, that's your issue."
        }
    }
}

# ----------------------------{ File Browser }---------------------------
function Show-FileBrowserDialog {
    param(
        [string]$StartDir = ".",
        [string]$Title = "Select File",
        [string[]]$Filter = @("*.*")
    )

    $script:selectedFile = $null  # Use script scope from the start
    $currentPath = (Resolve-Path $StartDir).Path
    $dialog = [Terminal.Gui.Dialog]::new($Title, 80, 24)

    # Current path label
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
            Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | Sort-Object Name |
            ForEach-Object { $items.Add("[DIR] $($_.Name)") }

            ## Files matching filter
            Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue |
            Where-Object { $Filter -contains "*.*" -or $Filter -contains "*$($_.Extension)" } |
            Sort-Object Name | ForEach-Object { $items.Add($_.Name) }
        } catch { Show-Modal "Error" "Cannot access directory: $path" }

        if ($items.Count -eq 0) { $items.Add("(empty directory)") }
        $listView.SetSource($items)
    }

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
    Update-FileList -path $currentPath

    ## Run dialog
    [Terminal.Gui.Application]::Run($dialog)

    return $script:selectedFile  # Return the script-scoped variable
}

function Initialize-UIFramework {
  <#
  .SYNOPSIS
  Initialize Terminal.Gui application and create main window

  .DESCRIPTION
  Sets up the Terminal.Gui framework, creates the top-level application
  and main window, and applies the selected theme. This should be called
  FIRST before any data loading to ensure the UI is visible.

  .PARAMETER Theme
  Theme name to apply (HighContrast, PanAm, Matrix, etc)

  .PARAMETER Title
  Window title to display

  .EXAMPLE
  $uiComponents = Initialize-UIFramework -Theme "PanAm" -Title "DSA-TUI v1.0"
  $top = $uiComponents.Top
  $win = $uiComponents.Window
  #>

  param(
    [string]$Theme = "HighContrast",
    [string]$Title = "DSA-TUI - Active Directory"
  )

  Debug-Log ": Initializing Terminal.Gui framework..." -Type "Info"

  ## ==================== STEP 1: Initialize Terminal.Gui ====================
  try {
    [Terminal.Gui.Application]::Init()
    Debug-Log ": Terminal.Gui.Application initialized" -Type "Success"
  } catch {
    Debug-Log ": FATAL - Failed to initialize Terminal.Gui: $($_.Exception.Message)" -Type "Error"
    throw
  }

  ## Get top-level application
  $top = [Terminal.Gui.Application]::Top

  ## ==================== STEP 2: Create Main Window ====================
  $win = [Terminal.Gui.Window]::new($Title)
  $win.X = 0
  $win.Y = 0
  $win.Width  = [Terminal.Gui.Dim]::Fill()
  $win.Height = [Terminal.Gui.Dim]::Fill(1)  # Leave room for status bar

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
    } else {
      Debug-Log ": WARNING - Theme data is null, using defaults" -Type "Warn"
    }
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
function Initialize-UIFramework {
  <#
  .SYNOPSIS
  Initialize Terminal.Gui application and create main window

  .DESCRIPTION
  Sets up the Terminal.Gui framework, creates the top-level application
  and main window, and applies the selected theme. This should be called
  FIRST before any data loading to ensure the UI is visible.

  .PARAMETER Theme
  Theme name to apply (HighContrast, PanAm, Matrix, etc)

  .PARAMETER Title
  Window title to display

  .EXAMPLE
  $uiComponents = Initialize-UIFramework -Theme "PanAm" -Title "DSA-TUI v1.0"
  $top = $uiComponents.Top
  $win = $uiComponents.Window
  #>

  param(
    [string]$Theme = "HighContrast",
    [string]$Title = "DSA-TUI - Active Directory"
  )

  Debug-Log ": Initializing Terminal.Gui framework..." -Type "Info"

  ## ==================== STEP 1: Initialize Terminal.Gui ====================
  try {
    [Terminal.Gui.Application]::Init()
    Debug-Log ": Terminal.Gui.Application initialized" -Type "Success"
  } catch {
    Debug-Log ": FATAL - Failed to initialize Terminal.Gui: $($_.Exception.Message)" -Type "Error"
    throw
  }

  ## Get top-level application
  $top = [Terminal.Gui.Application]::Top

  ## ==================== STEP 2: Create Main Window ====================
  $win = [Terminal.Gui.Window]::new($Title)
  $win.X = 0
  $win.Y = 0
  $win.Width  = [Terminal.Gui.Dim]::Fill()
  $win.Height = [Terminal.Gui.Dim]::Fill(1)  # Leave room for status bar

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
    } else {
      Debug-Log ": WARNING - Theme data is null, using defaults" -Type "Warn"
    }
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
      if ($Object.Enabled) {
        [void]$menuItems.Add("Disable Account")
      } else {
        [void]$menuItems.Add("Enable Account")
      }

      if ($Object.LockedOut -or $Object.Locked) {
        [void]$menuItems.Add("Unlock Account")
      }

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
  Universal status bar management - initialize, update, and animate

  .DESCRIPTION
  Single function that handles all status bar operations:
  - Initialize: Create status bar with F-key shortcuts and theme
  - Update: Set static messages
  - Spinner: Show animated spinner for long operations
  - Final: Mark operation complete with checkmark

  .PARAMETER Initialize
  Create and configure the status bar (call once at startup)

  .PARAMETER ThemeData
  Theme data to apply (used with -Initialize)

  .PARAMETER Message
  Status message to display (used for updates)

  .PARAMETER Spinner
  Show animated spinner (for long operations)

  .PARAMETER Final
  Mark operation as complete (shows checkmark)

  .EXAMPLE
  # Initialize at startup
  $statusBar = Set-StatusBar -Initialize -ThemeData $Script:themeData
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

    [switch]$Initialize,
    [object]$ThemeData = $null,
    [switch]$Spinner,
    [switch]$Final
  )

  ## ==================== INITIALIZE MODE ====================
  if ($Initialize) {
    Debug-Log ": Initializing status bar..." -Type "Info"

    ## Initialize spinner globals
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
      @{ Key = [Terminal.Gui.Key]::F1;  Label = "~F1~ Help";     Action = { Show-Modal "Shortcuts" "F1-Help | F2-Password | F3-New | F5-Refresh | F6-Themes | F7-Search | F10-Quit" } }
      @{ Key = [Terminal.Gui.Key]::F2;  Label = "~F2~ Password"; Action = { Generate-RandomPassword } }
      @{ Key = [Terminal.Gui.Key]::F3;  Label = "~F3~ New";      Action = { Show-NewObjectWizard } }
      @{ Key = [Terminal.Gui.Key]::F5;  Label = "~F5~ Refresh";  Action = { Refresh-Data -domain $Script:CurrentDomain -RebuildTree } }
      @{ Key = [Terminal.Gui.Key]::F6;  Label = "~F6~ Themes";   Action = { Show-ThemeSelector } }
      @{ Key = [Terminal.Gui.Key]::F7;  Label = "~F7~ Search";   Action = { Show-ADSearchDialog } }
      @{ Key = [Terminal.Gui.Key]::F9;  Label = "~F9~ Menus";    Action = { } }
      @{ Key = [Terminal.Gui.Key]::F10; Label = "~F10~ Quit";    Action = { [Terminal.Gui.Application]::RequestStop() } }
      @{ Key = [Terminal.Gui.Key]::F11; Label = "~F11~ Full";    Action = { } }
    )

    ## Build status items array
    $items = @()
    foreach ($sc in $shortcuts) {
      $items += [Terminal.Gui.StatusItem]::new($sc.Key, $sc.Label, $sc.Action)
    }
    $items += $Script:StatusItem

    ## Create status bar
    $Script:StatusBar = [Terminal.Gui.StatusBar]::new($items)

    ## Apply theme if provided
    if ($ThemeData -and $ThemeData.StatusBar) {
      $Script:StatusBar.ColorScheme = $ThemeData.StatusBar
      Debug-Log ": Status bar theme applied" -Type "Success"
    }

    Debug-Log ": Status bar initialized with $($shortcuts.Count) shortcuts" -Type "Success"
    return $Script:StatusBar
  }

  ## ==================== UPDATE MODE ====================

  ## Guard: Status bar must exist
  if (-not $Script:StatusItem -or -not $Script:StatusBar) {
    Debug-Log ": StatusBar not initialized - call 'Set-StatusBar -Initialize' first!" -Type "Error"
    Write-Warning "StatusBar not initialized. Call Set-StatusBar -Initialize before using."
    return
  }

  ## Build dynamic prefix (refreshed on each call)
  $prefix = "Forest: $($Script:ForestName) | Objs: $($Script:ADObjects.Count)"

  ## ==================== SPINNER MODE ====================
  if ($Spinner) {
    $Script:SpinnerActive = $true
    $Script:statusBaseMessage = $Message
    $Script:statusSpinnerIndex = 0
    $Script:statusPrefix = $prefix  # Cache for timer

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
    if ($Script:SpinnerTimer.Enabled) {
      $Script:SpinnerTimer.Stop()
    }
    $Script:SpinnerTimer.Start()

    Debug-Log ": Spinner started: $Message" -Type "Info"
    return
  }

  ## ==================== STATIC/FINAL MODE ====================

  ## Stop spinner if running
  $Script:SpinnerActive = $false
  if ($Script:SpinnerTimer) {
    $Script:SpinnerTimer.Stop()
  }

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

function Show-Modal {
  param(
    [string]$title,
    [string]$msg,
    [switch]$YesNo
  )

  if ($YesNo) {
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, $title, $msg, "Yes", "No")
    return $result  # Returns 0 for Yes, 1 for No
  }
  else {
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null
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

## -------------------------{ Get Theme }-------------------------
## ==================== UNIFIED THEME MANAGEMENT ====================
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
    if (-not $Mode) {
        throw "Get-Theme called with empty mode"
    }

    ## Initialize color schemes and Ensure ColorSchemes are instantiated
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

## -------------------------{ Apply Colours }-------------------------
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

## -------------------------{ Danske Soda vand }-------------------------
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


################### Demo Data Functions ################
function Manage-DemoData {
  <#
  .SYNOPSIS
  Unified demo data management - import, export, backup, and restore

  .DESCRIPTION
  Single function to handle all demo data operations with helper functions properly scoped
  #>

  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateSet('Import', 'Export', 'Backup', 'Restore', 'Reset')]
    [string]$Action,

    [Parameter(Position=1)]
    [string]$Path = "",

    [switch]$AsDemo,

    [bool]$IncludeUsers = $true,
    [bool]$IncludeGroups = $true,
    [bool]$IncludeComputers = $true,
    [bool]$IncludeDCs = $true
  )

  ## ==================== INTERNAL HELPER FUNCTIONS ====================
  ## Define these FIRST so they're available to all code below

  function ExtractOUPath {
    param([string]$DN)
    if ([string]::IsNullOrWhiteSpace($DN)) { return @() }
    $ouComponents = @()
    $parts = $DN -split '(?<!\\),' | Where-Object { $_ -match '^OU=' }
    foreach ($part in $parts) {
      if ($part -match '^OU=(.+)$') { $ouComponents += $matches[1] }
    }
    [array]::Reverse($ouComponents)
    return $ouComponents
  }

  function ExtractDomain {
    param([string]$DN)
    if ([string]::IsNullOrWhiteSpace($DN)) { return '' }
    $dcParts = $DN -split '(?<!\\),' | Where-Object { $_ -match '^DC=' } | ForEach-Object {
      if ($_ -match '^DC=(.+)$') { $matches[1] }
    }
    return ($dcParts -join '.')
  }

  function ParseUser {
    param([object]$CSVObject)
    try {
      $dn = $CSVObject.distinguishedName
      $ouPath = ExtractOUPath -DN $dn
      $domain = ExtractDomain -DN $dn
      $groups = @()
      if ($CSVObject.memberOf) {
        $groups = $CSVObject.memberOf -split ';' | ForEach-Object {
          if ($_ -match 'CN=([^,]+)') { $matches[1] }
        }
      }
      $uac = [int]$CSVObject.userAccountControl
      $disabled = ($uac -band 0x0002) -ne 0
      $locked = ($uac -band 0x0010) -ne 0
      return @{
        Name = $CSVObject.name; SamAccountName = $CSVObject.sAMAccountName
        UserPrincipalName = $CSVObject.userPrincipalName; OU = $ouPath; Groups = $groups
        Title = $CSVObject.title; Email = $CSVObject.mail; Country = $CSVObject.countryCode
        Disabled = $disabled; Locked = $locked
        MustChangePassword = ($CSVObject.pwdLastSet -eq '0')
        Department = $CSVObject.description; Office = $CSVObject.physicalDeliveryOfficeName
        Phone = $CSVObject.telephoneNumber; MobilePhone = $CSVObject.mobile
        Description = $CSVObject.description; Manager = ''; Company = ''; Domain = $domain
      }
    } catch {
      Debug-Log ": Failed to parse user: $($CSVObject.name)" -Type "Warn"
      return $null
    }
  }

  function ParseGroup {
    param([object]$CSVObject)
    try {
      $groupType = [int]$CSVObject.groupType
      $isSecurity = ($groupType -band 0x80000000) -ne 0
      $type = if ($isSecurity) { 'Security' } else { 'Distribution' }
      $scope = switch ($groupType -band 0xF) {
        1 { 'Global' }; 2 { 'DomainLocal' }; 4 { 'Universal' }; default { 'Global' }
      }
      return @{
        Name = $CSVObject.name; Description = $CSVObject.description
        Type = $type; Scope = $scope
        ManagedBy = if ($CSVObject.managedBy -match 'CN=([^,]+)') { $matches[1] } else { '' }
        Email = $CSVObject.mail
      }
    } catch { return $null }
  }

  function ParseComputer {
    param([object]$CSVObject)
    try {
      $dn = $CSVObject.distinguishedName
      $ouPath = ExtractOUPath -DN $dn
      $domain = ExtractDomain -DN $dn
      $uac = [int]$CSVObject.userAccountControl
      $enabled = ($uac -band 0x0002) -eq 0
      return @{
        Name = $CSVObject.name; SamAccountName = $CSVObject.sAMAccountName
        Type = 'Computer'; Role = 'WKS'; OU = $ouPath
        OS = $CSVObject.operatingSystem; OperatingSystemVersion = $CSVObject.operatingSystemVersion
        DNSHostName = $CSVObject.dNSHostName; Description = $CSVObject.description
        Enabled = $enabled; Domain = $domain
      }
    } catch { return $null }
  }

  function ParseDomainController {
    param([object]$CSVObject)
    try {
      $domain = ExtractDomain -DN $CSVObject.distinguishedName
      $isGC = $CSVObject.servicePrincipalName -match 'GC/'
      return @{
        Name = $CSVObject.name; SamAccountName = $CSVObject.sAMAccountName
        DNSHostName = $CSVObject.dNSHostName; Site = 'Default-First-Site-Name'
        Location = ''; Domain = $domain; Forest = $domain
        OS = $CSVObject.operatingSystem; OperatingSystemVersion = $CSVObject.operatingSystemVersion
        IPv4Address = ''; Enabled = $true; IsGlobalCatalog = $isGC
        FSMORoles = @(); LastReplication = Get-Date
        ReplicationHealth = 'Healthy'; LastBoot = Get-Date
      }
    } catch { return $null }
  }

  function ParseOU {
    param([object]$CSVObject)
    try {
      $dn = $CSVObject.distinguishedName
      $ouPath = ExtractOUPath -DN $dn
      return @{
        Name = $CSVObject.ou; DN = $dn; Path = $ouPath
        Description = $CSVObject.description
      }
    } catch { return $null }
  }

  function BuildUserCSV {
    param([hashtable]$User)
    $ouDN = ($User.OU | ForEach-Object { "OU=$_" }) -join ','
    $dn = "CN=$($User.Name),$ouDN,DC=example,DC=com"
    $uac = 512
    if ($User.Disabled) { $uac += 2 }
    if ($User.Locked) { $uac += 16 }
    return [PSCustomObject]@{
      distinguishedName = $dn; objectClass = 'user'; name = $User.Name
      sAMAccountName = $User.SamAccountName; userPrincipalName = $User.UserPrincipalName
      mail = $User.Email; title = $User.Title; description = $User.Description
      telephoneNumber = $User.Phone; mobile = $User.MobilePhone
      userAccountControl = $uac
      pwdLastSet = if ($User.MustChangePassword) { '0' } else { '132000000000000000' }
      memberOf = ($User.Groups | ForEach-Object { "CN=$_,OU=Groups,DC=example,DC=com" }) -join ';'
    }
  }

  function BuildGroupCSV {
    param([hashtable]$Group)
    $groupType = if ($Group.Type -eq 'Security') { -2147483646 } else { 2 }
    return [PSCustomObject]@{
      distinguishedName = "CN=$($Group.Name),OU=Groups,DC=example,DC=com"
      objectClass = 'group'; name = $Group.Name; description = $Group.Description
      mail = $Group.Email; groupType = $groupType
    }
  }

  function BuildComputerCSV {
    param([hashtable]$Computer)
    $ouDN = ($Computer.OU | ForEach-Object { "OU=$_" }) -join ','
    $dn = "CN=$($Computer.Name),$ouDN,DC=example,DC=com"
    $uac = if ($Computer.Enabled) { 4096 } else { 4098 }
    return [PSCustomObject]@{
      distinguishedName = $dn; objectClass = 'computer'; name = $Computer.Name
      sAMAccountName = $Computer.SamAccountName; dNSHostName = $Computer.DNSHostName
      operatingSystem = $Computer.OS; operatingSystemVersion = $Computer.OperatingSystemVersion
      description = $Computer.Description; userAccountControl = $uac
    }
  }

  function BuildDCCSV {
    param([hashtable]$DC)
    $dn = "CN=$($DC.Name),OU=Domain Controllers,DC=example,DC=com"
    $spn = "HOST/$($DC.DNSHostName)"
    if ($DC.IsGlobalCatalog) { $spn += ";GC/$($DC.DNSHostName)" }
    return [PSCustomObject]@{
      distinguishedName = $dn; objectClass = 'computer'; name = $DC.Name
      sAMAccountName = $DC.SamAccountName; dNSHostName = $DC.DNSHostName
      operatingSystem = $DC.OS; operatingSystemVersion = $DC.OperatingSystemVersion
      userAccountControl = 532480; servicePrincipalName = $spn
    }
  }

  ## ==================== IMPORT ====================
  if ($Action -eq 'Import') {
    if ([string]::IsNullOrWhiteSpace($Path)) {
      $Path = Show-FileBrowserDialog -StartDir $PSScriptRoot -Title "Import AD CSV Data" -Filter @("*.csv", "*.txt")
      if (-not $Path) { Debug-Log ": Import cancelled" -Type "Info"; return }
    }
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }

    Debug-Log ": Importing AD data from CSV: $Path" -Type "Info"
    try {
      $csvData = Import-Csv -Path $Path -ErrorAction Stop
      Debug-Log ": Loaded $($csvData.Count) objects from CSV" -Type "Info"

      $users = [System.Collections.ArrayList]@()
      $groups = [System.Collections.ArrayList]@()
      $computers = [System.Collections.ArrayList]@()
      $dcs = [System.Collections.ArrayList]@()
      $ous = [System.Collections.ArrayList]@()

      ## ---- Infrastructure / non-demo AD objects ----
      $trusts        = [System.Collections.ArrayList]@()
      $servicePoints = [System.Collections.ArrayList]@()
      $dfsrObjects   = [System.Collections.ArrayList]@()
      $unknownInfra  = [System.Collections.ArrayList]@()

      foreach ($obj in $csvData) {
        $objectClass = $obj.objectClass
        switch -Regex ($objectClass) {
          'user' { $user = ParseUser -CSVObject $obj; if ($user) { [void]$users.Add($user) } }
          'group' { $group = ParseGroup -CSVObject $obj; if ($group) { [void]$groups.Add($group) } }
          'computer' {
            if ($obj.userAccountControl -match '532480|8192' -or $obj.servicePrincipalName -match 'E3514235') {
              $dc = ParseDomainController -CSVObject $obj; if ($dc) { [void]$dcs.Add($dc) }
            } else {
              $computer = ParseComputer -CSVObject $obj; if ($computer) { [void]$computers.Add($computer) }
            }
          }
          'organizationalUnit'     { $ou    = ParseOU -CSVObject $obj            ; if ($ou) { [void]$ous.Add($ou) } }
          'trustedDomain'          { $trust = ParseTrustedDomain -CSVObject $obj ; if ($trust) { [void]$trusts.Add($trust) }}
          'serviceConnectionPoint' { $scp   = ParseServiceConnectionPoint -CSVObject $obj ; if ($scp) { [void]$servicePoints.Add($scp) }}
          '^msDFSR-'               { $dfsr = ParseDFSRObject -CSVObject $obj     ;   if ($dfsr) { [void]$dfsrObjects.Add($dfsr) }}
          '^(domainDNS|container|builtinDomain)$' {}  # Structural AD containers – ignored but expected
          default {
            [void]$unknownInfra.Add(@{
            Class = $objectClass
            Name  = $obj.name
            DN    = $obj.distinguishedName
            })
          }
      }

      Debug-Log ": Parsed $($users.Count) users, $($groups.Count) groups, $($computers.Count) computers, $($dcs.Count) DCs" -Type "Success"

      if ($AsDemo) {
        $Script:rawUsers = $users; $Script:rawDCs = $dcs; $Script:rawComputers = $computers
        $Script:rawDemoGroups    = $groups; $Script:rawOUs = $ous
        $Script:rawTrusts        = $trusts
        $Script:rawServicePoints = $servicePoints
        $Script:rawDFSR          = $dfsrObjects
        $Script:rawUnknownInfra  = $unknownInfra

        Debug-Log ": Demo data replaced with imported CSV data" -Type "Success"
      }

      return @{ Users = $users; Groups = $groups; Computers = $computers; DCs = $dcs; OUs = $ous; Trusts = $trusts; ServicePoints = $servicePoints; DFSR = $dfsrObjects; UnknownInfra = $unknownInfra }
    }
  } catch {
    Debug-Log ": Failed to import CSV: $($_.Exception.Message)" -Type "Error"
    throw
  }

  ## ==================== EXPORT ====================
  elseif ($Action -eq 'Export') {
    if ([string]::IsNullOrWhiteSpace($Path)) {
      $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
      $Path = Show-FileBrowserDialog -StartDir $PSScriptRoot -Title "Export Demo Data" -Filter @("*.csv", "*.txt")
      if (-not $Path) { Debug-Log ": Export cancelled" -Type "Info"; return }
      if ($Path -notmatch '\.(csv|txt)$') { $Path = "$Path-$timestamp.csv" }
    }

    Debug-Log ": Exporting demo data to CSV: $Path" -Type "Info"
    try {
      $exportData = [System.Collections.ArrayList]@()
      if ($IncludeUsers -and $Script:rawUsers) {
        foreach ($user in $Script:rawUsers) { [void]$exportData.Add((BuildUserCSV -User $user)) }
        Debug-Log ": Added $($Script:rawUsers.Count) users to export" -Type "Info"
      }
      if ($IncludeGroups -and $Script:rawDemoGroups) {
        foreach ($group in $Script:rawDemoGroups) { [void]$exportData.Add((BuildGroupCSV -Group $group)) }
        Debug-Log ": Added $($Script:rawDemoGroups.Count) groups to export" -Type "Info"
      }
      if ($IncludeComputers -and $Script:rawComputers) {
        foreach ($computer in $Script:rawComputers) { [void]$exportData.Add((BuildComputerCSV -Computer $computer)) }
        Debug-Log ": Added $($Script:rawComputers.Count) computers to export" -Type "Info"
      }
      if ($IncludeDCs -and $Script:rawDCs) {
        foreach ($dc in $Script:rawDCs) { [void]$exportData.Add((BuildDCCSV -DC $dc)) }
        Debug-Log ": Added $($Script:rawDCs.Count) DCs to export" -Type "Info"
      }
      $exportData | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
      Debug-Log ": Exported $($exportData.Count) objects to $Path" -Type "Success"
      return $Path
    } catch {
      Debug-Log ": Failed to export CSV: $($_.Exception.Message)" -Type "Error"
      throw
    }
  }

  ## ==================== BACKUP ====================
  elseif ($Action -eq 'Backup') {
    $backupDir = Join-Path $PSScriptRoot "demo-backups"
    if (-not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force | Out-Null }
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = Join-Path $backupDir "demo-backup-$timestamp.csv"
    Debug-Log ": Creating backup at $backupPath" -Type "Info"
    Manage-DemoData -Action Export -Path $backupPath
    return $backupPath
  }

  ## ==================== RESTORE ====================
  elseif ($Action -eq 'Restore') {
    if ([string]::IsNullOrWhiteSpace($Path)) {
      $backupDir = Join-Path $PSScriptRoot "demo-backups"
      if (Test-Path $backupDir) {
        $latestBackup = Get-ChildItem -Path $backupDir -Filter "demo-backup-*.csv" |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestBackup) {
          $Path = $latestBackup.FullName
          Debug-Log ": Using most recent backup: $Path" -Type "Info"
        } else { throw "No backup files found in $backupDir" }
      } else { throw "Backup directory not found: $backupDir" }
    }
    Debug-Log ": Restoring demo data from backup: $Path" -Type "Info"
    Manage-DemoData -Action Import -Path $Path -AsDemo
  }

  ## ==================== RESET ====================
  elseif ($Action -eq 'Reset') {
    Debug-Log ": Resetting demo data to hardcoded defaults" -Type "Info"
    $Script:rawUsers = $null; $Script:rawDCs = $null; $Script:rawComputers = $null
    $Script:rawDemoGroups = $null; $Script:rawOUs = $null
    Debug-Log ": Demo data cleared - will reload defaults on next refresh" -Type "Success"
  }

  function ParseTrustedDomain {
    param([object]$CSVObject)
    return @{
      Name      = $CSVObject.name
      Partner  = $CSVObject.trustPartner
      Type     = $CSVObject.trustType
      Direction= $CSVObject.trustDirection
      DN       = $CSVObject.distinguishedName
    }
  }

  function ParseServiceConnectionPoint {
    param([object]$CSVObject)
    return @{
      Name     = $CSVObject.name
      Service  = $CSVObject.serviceClassName
      Keywords = if ($CSVObject.keywords) { $CSVObject.keywords -split ';' } else { @() }
      DN       = $CSVObject.distinguishedName
    }
  }

  function ParseDFSRObject {
    param([object]$CSVObject)
    return @{
      Class = $CSVObject.objectClass
      Name  = $CSVObject.name
      DN    = $CSVObject.distinguishedName
    }
  }
  }

  Debug-Log ": Infra: $($trusts.Count) trusts, $($servicePoints.Count) services, $($dfsrObjects.Count) DFSR objects" -Type "Debug"
}

############################# AD Funcitons #######################

## Called once during initialization - If we have Terminal-Icons pwsh module
## and a suitable nerd font, use icons instead of text
function Show-UserPropertiesDialog {
  param($user)

  ## ---------------------- Safety Checks ----------------------
  if (-not $user) {
    Debug-Log ": User object is null" -Type "Warn"
    return
  }
  Debug-Log ": Show-UserPropertiesDialog starting for: $($user.Name)" -Type "Info"

  ## ==================== Tab Definitions ====================

  # General Tab
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $user, $state)
      $y = 1

      ## Basic Info
      $lbl = [Terminal.Gui.Label]::new("User Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)

      $state.txtName = [Terminal.Gui.Label]::new($user.Name ?? ""); $state.txtName.X=20; $state.txtName.Y=$y
      $view.Add($state.txtName); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDisplayName = [Terminal.Gui.TextField]::new($user.DisplayName ?? ""); $state.txtDisplayName.X=20; $state.txtDisplayName.Y=$y; $state.txtDisplayName.Width=60
      $view.Add($state.txtDisplayName); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Email:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $emailAddr = if ($user.EmailAddress) { $user.EmailAddress } elseif ($user.mail) { $user.mail } else { "" }
      $state.txtEmail = [Terminal.Gui.TextField]::new($emailAddr); $state.txtEmail.X=20; $state.txtEmail.Y=$y; $state.txtEmail.Width=60
      $view.Add($state.txtEmail); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDescription = [Terminal.Gui.TextField]::new($user.Description ?? ""); $state.txtDescription.X=20; $state.txtDescription.Y=$y; $state.txtDescription.Width=60
      $view.Add($state.txtDescription); $y+=2

      ## Phone numbers
      $lbl = [Terminal.Gui.Label]::new("Contact Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtOfficePhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $state.txtOfficePhone.X=20; $state.txtOfficePhone.Y=$y; $state.txtOfficePhone.Width=30
      $view.Add($state.txtOfficePhone); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Mobile:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtMobilePhone = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $state.txtMobilePhone.X=20; $state.txtMobilePhone.Y=$y; $state.txtMobilePhone.Width=30
      $view.Add($state.txtMobilePhone); $y+=1
    }
  }

  # Account Tab
  $accountTab = @{
    Name = "Account"
    Builder = {
      param($view, $user, $state)
      $y = 1

      ## Logon Information at the TOP
      $lbl = [Terminal.Gui.Label]::new("Logon Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $lbl = [Terminal.Gui.Label]::new("User logon name (UPN):"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $upn = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "" }
      $state.txtUserPrincipalName = [Terminal.Gui.TextField]::new($upn)
      $state.txtUserPrincipalName.X=35; $state.txtUserPrincipalName.Y=$y; $state.txtUserPrincipalName.Width=50
      $view.Add($state.txtUserPrincipalName); $y+=1

      $lbl = [Terminal.Gui.Label]::new("User logon name (Pre Win 2000):"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtSamAccountName = [Terminal.Gui.TextField]::new($user.SamAccountName ?? "")
      $state.txtSamAccountName.X=35; $state.txtSamAccountName.Y=$y; $state.txtSamAccountName.Width=50
      $view.Add($state.txtSamAccountName); $y+=2

      ## Account Status
      $lbl = [Terminal.Gui.Label]::new("Account Status"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $isEnabled = if ($user.PSObject.Properties['Enabled']) { $user.Enabled } else { -not $user.Disabled }
      $state.chkEnabled = [Terminal.Gui.CheckBox]::new("Account Enabled"); $state.chkEnabled.X=4; $state.chkEnabled.Y=$y; $state.chkEnabled.Checked=$isEnabled
      $view.Add($state.chkEnabled); $y+=1

      $isLocked = if ($user.PSObject.Properties['LockedOut']) { $user.LockedOut } else { $user.Locked ?? $false }
      $state.chkLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $state.chkLocked.X=4; $state.chkLocked.Y=$y; $state.chkLocked.Checked=$isLocked; $state.chkLocked.Enabled=$false
      $view.Add($state.chkLocked); $y+=2

      ## ==================== ACCOUNT EXPIRATION ADDITION TO USER PROPERTIES ====================
##
## ADD THIS CODE TO THE GENERAL TAB BUILDER IN Show-UserPropertiesDialog
## Insert after the "Enabled" checkbox section, before closing the General tab
##

## Account Expiration Section
$y += 2
$lbl = [Terminal.Gui.Label]::new("Account Expiration"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

## Never Expires Checkbox
$state.chkNeverExpires = [Terminal.Gui.CheckBox]::new("Never Expires")
$state.chkNeverExpires.X = 4
$state.chkNeverExpires.Y = $y

## Get current expiration status
$hasExpiration = $false
$expirationDate = $null

if ($Script:DemoMode) {
    ## Demo mode - check if user has AccountExpirationDate property
    if ($user.PSObject.Properties.Match('AccountExpirationDate') -and $user.AccountExpirationDate) {
        $hasExpiration = $true
        $expirationDate = $user.AccountExpirationDate
    }
} else {
    ## Production mode - query AD
    try {
        $adUser = Get-ADUser -Identity $user.SamAccountName -Properties AccountExpirationDate -ErrorAction Stop
        if ($adUser.AccountExpirationDate) {
            $hasExpiration = $true
            $expirationDate = $adUser.AccountExpirationDate
        }
    } catch {
        Debug-Log ": Failed to get account expiration: $($_.Exception.Message)" -Type "Warn"
    }
}

$state.chkNeverExpires.Checked = -not $hasExpiration
$view.Add($state.chkNeverExpires)
$y += 1

## Expiration Date Field
$lbl = [Terminal.Gui.Label]::new("Expires On:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)

$expirationDateStr = if ($expirationDate) {
    $expirationDate.ToString("yyyy-MM-dd")
} else {
    ""
}

$state.txtExpirationDate = [Terminal.Gui.TextField]::new($expirationDateStr)
$state.txtExpirationDate.X = 20
$state.txtExpirationDate.Y = $y
$state.txtExpirationDate.Width = 20
$state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
$view.Add($state.txtExpirationDate)

## Format hint
$lblFormat = [Terminal.Gui.Label]::new("(yyyy-MM-dd)")
$lblFormat.X = 42
$lblFormat.Y = $y
$view.Add($lblFormat)
$y += 1

## Days Until Expiry (calculated, read-only)
$lbl = [Terminal.Gui.Label]::new("Days Until Expiry:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)

$daysUntilExpiry = if ($expirationDate) {
    $days = ($expirationDate - (Get-Date)).Days
    if ($days -lt 0) {
        "EXPIRED ($([Math]::Abs($days)) days ago)"
    } elseif ($days -eq 0) {
        "TODAY"
    } else {
        "$days days"
    }
} else {
    "N/A"
}

$state.lblDaysUntilExpiry = [Terminal.Gui.Label]::new($daysUntilExpiry)
$state.lblDaysUntilExpiry.X = 20
$state.lblDaysUntilExpiry.Y = $y
$view.Add($state.lblDaysUntilExpiry)

## Toggle date field when checkbox changes
$state.chkNeverExpires.add_Toggled({
    $state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
    if ($state.chkNeverExpires.Checked) {
        $state.txtExpirationDate.Text = [NStack.ustring]::Make("")
        $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("N/A")
    }
}.GetNewClosure())

## Update days calculation when date changes
$state.txtExpirationDate.add_TextChanged({
    $dateStr = $state.txtExpirationDate.Text.ToString()
    if (-not [string]::IsNullOrWhiteSpace($dateStr)) {
        try {
            $parsedDate = [DateTime]::Parse($dateStr)
            $days = ($parsedDate - (Get-Date)).Days
            if ($days -lt 0) {
                $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("EXPIRED ($([Math]::Abs($days)) days ago)")
            } elseif ($days -eq 0) {
                $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("TODAY")
            } else {
                $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("$days days")
            }
        } catch {
            $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid date")
        }
    } else {
        $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("N/A")
    }
}.GetNewClosure())


## ==================== APPLY LOGIC ADDITION ====================
##
## ADD THIS TO THE Apply-ObjectChanges FUNCTION (in the User section)
## OR add to the existing Apply logic in Show-UserPropertiesDialog
##

## Handle Account Expiration
if ($state.chkNeverExpires -and $state.txtExpirationDate) {
    $neverExpires = $state.chkNeverExpires.Checked
    $expirationDateStr = $state.txtExpirationDate.Text.ToString().Trim()

    if ($neverExpires) {
        ## Clear expiration
        if ($Script:DemoMode) {
            if ($user.PSObject.Properties.Match('AccountExpirationDate')) {
                $user.AccountExpirationDate = $null
            }
            Debug-Log ":   Cleared account expiration (demo mode)" -Type "Info"
        } else {
            Clear-ADAccountExpiration -Identity $user.SamAccountName -ErrorAction Stop
            Debug-Log ":   Cleared account expiration in AD" -Type "Success"
        }
    } else {
        ## Set expiration date
        if (-not [string]::IsNullOrWhiteSpace($expirationDateStr)) {
            try {
                $expirationDate = [DateTime]::Parse($expirationDateStr)

                if ($Script:DemoMode) {
                    if ($user.PSObject.Properties.Match('AccountExpirationDate')) {
                        $user.AccountExpirationDate = $expirationDate
                    } else {
                        $user | Add-Member -NotePropertyName 'AccountExpirationDate' -NotePropertyValue $expirationDate -Force
                    }
                    Debug-Log ":   Set account expiration to $($expirationDate.ToString('yyyy-MM-dd')) (demo mode)" -Type "Info"
                } else {
                    Set-ADAccountExpiration -Identity $user.SamAccountName -DateTime $expirationDate -ErrorAction Stop
                    Debug-Log ":   Set account expiration to $($expirationDate.ToString('yyyy-MM-dd')) in AD" -Type "Success"
                }
            } catch {
                Debug-Log ":   Failed to parse/set expiration date: $($_.Exception.Message)" -Type "Error"
                throw "Invalid expiration date format"
            }
        }
    }
}




      ## Password Settings
      $lbl = [Terminal.Gui.Label]::new("Password Settings"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired"); $state.chkPasswordExpired.X=4; $state.chkPasswordExpired.Y=$y; $state.chkPasswordExpired.Checked=($user.PasswordExpired??$false); $state.chkPasswordExpired.Enabled=$false
      $view.Add($state.chkPasswordExpired); $y+=1

      $state.chkMustChangePassword = [Terminal.Gui.CheckBox]::new("User must change password at next logon"); $state.chkMustChangePassword.X=4; $state.chkMustChangePassword.Y=$y
      $state.chkMustChangePassword.Checked = if ($user.PasswordNeverExpires){$false}else{ if ($user.PSObject.Properties['pwdLastSet']){ $user.pwdLastSet -eq 0 } else { $false } }
      $view.Add($state.chkMustChangePassword); $y+=1

      $state.chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("User cannot change password"); $state.chkCannotChangePassword.X=4; $state.chkCannotChangePassword.Y=$y; $state.chkCannotChangePassword.Checked=($user.CannotChangePassword??$false)
      $view.Add($state.chkCannotChangePassword); $y+=1

      $state.chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires"); $state.chkPasswordNeverExpires.X=4; $state.chkPasswordNeverExpires.Y=$y; $state.chkPasswordNeverExpires.Checked=($user.PasswordNeverExpires??$false)
      $view.Add($state.chkPasswordNeverExpires); $y+=2

      ## Logon History
      $lbl = [Terminal.Gui.Label]::new("Logon History"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Last logon: "+($user.LastLogonDate?.ToString('yyyy-MM-dd HH:mm') ?? 'Never')); $lbl.X=4; $lbl.Y=$y
      $view.Add($lbl); $y+=1

      if ($user.PSObject.Properties['PasswordLastSet'] -and $user.PasswordLastSet) {
        $lbl = [Terminal.Gui.Label]::new("Password last set: "+$user.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($user.PSObject.Properties['LogonCount'] -or $user.PSObject.Properties['logonCount']) {
        $logonCount = if ($user.LogonCount) { $user.LogonCount } else { $user.logonCount }
        $lbl = [Terminal.Gui.Label]::new("Logon count: $logonCount"); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }
    }
  }

  # Address Tab
  $addressTab = @{
    Name = "Address"
    Builder = {
      param($view, $user, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Street:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtStreet = [Terminal.Gui.TextField]::new($user.StreetAddress ?? ""); $state.txtStreet.X=20; $state.txtStreet.Y=$y; $state.txtStreet.Width=70
      $view.Add($state.txtStreet); $y+=2

      $lbl = [Terminal.Gui.Label]::new("City:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtCity = [Terminal.Gui.TextField]::new($user.City ?? ""); $state.txtCity.X=20; $state.txtCity.Y=$y; $state.txtCity.Width=70
      $view.Add($state.txtCity); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Postal Code:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtPostal = [Terminal.Gui.TextField]::new($user.PostalCode ?? ""); $state.txtPostal.X=20; $state.txtPostal.Y=$y; $state.txtPostal.Width=20
      $view.Add($state.txtPostal); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Country:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtCountry = [Terminal.Gui.TextField]::new($user.Country ?? ""); $state.txtCountry.X=20; $state.txtCountry.Y=$y; $state.txtCountry.Width=70
      $view.Add($state.txtCountry); $y+=1
    }
  }

  # Profile Tab
  $profileTab = @{
    Name = "Profile"
    Builder = {
      param($view, $user, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("User Profile"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Profile path:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtProfilePath = [Terminal.Gui.TextField]::new($user.ProfilePath ?? ""); $state.txtProfilePath.X=20; $state.txtProfilePath.Y=$y; $state.txtProfilePath.Width=70
      $view.Add($state.txtProfilePath); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Logon script:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtLogonScript = [Terminal.Gui.TextField]::new($user.ScriptPath ?? ""); $state.txtLogonScript.X=20; $state.txtLogonScript.Y=$y; $state.txtLogonScript.Width=70
      $view.Add($state.txtLogonScript); $y+=3

      ## Home Folder
      $lbl = [Terminal.Gui.Label]::new("Home Folder"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Home directory:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtHomeDirectory = [Terminal.Gui.TextField]::new($user.HomeDirectory ?? ""); $state.txtHomeDirectory.X=20; $state.txtHomeDirectory.Y=$y; $state.txtHomeDirectory.Width=70
      $view.Add($state.txtHomeDirectory); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Home drive:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtHomeDrive = [Terminal.Gui.TextField]::new($user.HomeDrive ?? ""); $state.txtHomeDrive.X=20; $state.txtHomeDrive.Y=$y; $state.txtHomeDrive.Width=5
      $view.Add($state.txtHomeDrive); $y+=1
    }
  }

  # Organization Tab
  $organizationTab = @{
    Name = "Organization"
    Builder = {
      param($view, $user, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Title:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtTitle = [Terminal.Gui.TextField]::new($user.Title ?? ""); $state.txtTitle.X=20; $state.txtTitle.Y=$y; $state.txtTitle.Width=70
      $view.Add($state.txtTitle); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Department:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDept = [Terminal.Gui.TextField]::new($user.Department ?? ""); $state.txtDept.X=20; $state.txtDept.Y=$y; $state.txtDept.Width=70
      $view.Add($state.txtDept); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Company:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtCompany = [Terminal.Gui.TextField]::new($user.Company ?? ""); $state.txtCompany.X=20; $state.txtCompany.Y=$y; $state.txtCompany.Width=70
      $view.Add($state.txtCompany); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Manager:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtManager = [Terminal.Gui.TextField]::new($user.Manager ?? ""); $state.txtManager.X=20; $state.txtManager.Y=$y; $state.txtManager.Width=70
      $view.Add($state.txtManager); $y+=1
    }
  }

  # Member Of Tab
  $memberOfTab = @{
    Name = "Member Of"
    Builder = {
      param($view, $user, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Group Memberships:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      ## Create ListView for groups
      $state.lstGroups = [Terminal.Gui.ListView]::new()
      $state.lstGroups.X = 2
      $state.lstGroups.Y = $y
      $state.lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstGroups.Height = 25

      ## Get group list
      $state.groupList = @()
      if ($user.Groups) {
        $state.groupList = $user.Groups
      } elseif ($user.MemberOf) {
        $state.groupList = $user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } }
      }

      if ($state.groupList.Count -gt 0) {
        $state.lstGroups.SetSource($state.groupList)
      } else {
        $state.lstGroups.SetSource(@("(No group memberships)"))
      }

      $view.Add($state.lstGroups)

      ## Add to Group button
      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2
      $btnAdd.Y = 28
      $btnAdd.add_Clicked({
        Show-EditGroupMembershipDialog -User $user -OnUpdate {
          Debug-Log ": Refreshing group list after add" -Type "Info"

          $refreshedGroups = @()
          if ($user.Groups) {
            $refreshedGroups = $user.Groups
          } elseif ($user.MemberOf) {
            $refreshedGroups = $user.MemberOf | ForEach-Object {
              if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
            }
          }

          if ($refreshedGroups.Count -gt 0) {
            $state.lstGroups.SetSource($refreshedGroups)
          } else {
            $state.lstGroups.SetSource(@("(No group memberships)"))
          }

          $state.groupList = $refreshedGroups
          Debug-Log ": Group list refreshed, now showing $($refreshedGroups.Count) groups" -Type "Success"
        }
      }.GetNewClosure())
      $view.Add($btnAdd)

      ## Remove from Group button
      $btnRemove = [Terminal.Gui.Button]::new("Remove from Group")
      $btnRemove.X = 22
      $btnRemove.Y = 28
      $btnRemove.add_Clicked({
        $selectedIndex = $state.lstGroups.SelectedItem

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

          if ($selectedGroup -eq "(No group memberships)") {
            Show-Modal "Info" "No group selected"
            return
          }

          $confirmDlg = Show-Modal "Confirm Removal" "Remove $($user.Name) from group '$selectedGroup'?" -YesNo

          if ($confirmDlg -eq 0) {
            try {
              if ($Script:DemoMode) {
                $user.Groups = $user.Groups | Where-Object { $_ -ne $selectedGroup }
                Debug-Log ": Removed $($user.Name) from group $selectedGroup (demo mode)" -Type "Success"
              } else {
                Remove-ADGroupMember -Identity $selectedGroup -Members $user.SamAccountName -Confirm:$false
                Debug-Log ": Removed $($user.Name) from group $selectedGroup" -Type "Success"
              }

              $updatedGroups = @()
              if ($user.Groups) {
                $updatedGroups = $user.Groups
              } elseif ($user.MemberOf) {
                $updatedGroups = $user.MemberOf | Where-Object { $_ -ne $selectedGroup } | ForEach-Object {
                  if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                }
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
              Debug-Log ": Failed to remove from group: $($_.Exception.Message)" -Type "Error"
            }
          }
        } else {
          Show-Modal "Info" "Please select a group to remove"
        }
      }.GetNewClosure())
      $view.Add($btnRemove)
    }
  }

  # Search/Lookup Tab
  $searchTab = @{
    Name = "Search/Lookup"
    Builder = {
      param($view, $user, $state)
      $y = 1

      $lblSearchDomain = [Terminal.Gui.Label]::new("Domain:"); $lblSearchDomain.X=2; $lblSearchDomain.Y=$y; $view.Add($lblSearchDomain)
      $state.txtSearchDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain ?? ""); $state.txtSearchDomain.X=15; $state.txtSearchDomain.Y=$y; $state.txtSearchDomain.Width=30
      $view.Add($state.txtSearchDomain)
      $y+=2

      $lblSearchName = [Terminal.Gui.Label]::new("Name:"); $lblSearchName.X=2; $lblSearchName.Y=$y; $view.Add($lblSearchName)
      $state.txtSearchUser = [Terminal.Gui.TextField]::new($user.Name ?? ""); $state.txtSearchUser.X=15; $state.txtSearchUser.Y=$y; $state.txtSearchUser.Width=30
      $view.Add($state.txtSearchUser)
      $y+=2

      $lblSearchType = [Terminal.Gui.Label]::new("Type:"); $lblSearchType.X=2; $lblSearchType.Y=$y; $view.Add($lblSearchType)
      $state.cmbSearchType = [Terminal.Gui.ComboBox]::new(); $state.cmbSearchType.X=15; $state.cmbSearchType.Y=$y; $state.cmbSearchType.Width=20
      $state.cmbSearchType.SetSource(@("User","Group","OU")); $state.cmbSearchType.SelectedItem = 0
      $view.Add($state.cmbSearchType)
      $y+=2

      $lblSearchFilter = [Terminal.Gui.Label]::new("Filter Results:"); $lblSearchFilter.X=48; $lblSearchFilter.Y=1; $view.Add($lblSearchFilter)
      $state.txtSearchFilter = [Terminal.Gui.TextField]::new(""); $state.txtSearchFilter.X=62; $state.txtSearchFilter.Y=1; $state.txtSearchFilter.Width=20
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

      $lblSearchResult = [Terminal.Gui.Label]::new("Results:"); $lblSearchResult.X=2; $lblSearchResult.Y=$y; $view.Add($lblSearchResult)
      $y+=1
      $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
      $state.txtSearchOutput.X=2; $state.txtSearchOutput.Y=$y
      $state.txtSearchOutput.Width=[Terminal.Gui.Dim]::Fill(2)
      $state.txtSearchOutput.Height=[Terminal.Gui.Dim]::Fill(4)
      $state.txtSearchOutput.ReadOnly=$true
      $state.txtSearchOutput.WordWrap=$false
      $view.Add($state.txtSearchOutput)

      $state.chkSearchLocked = [Terminal.Gui.CheckBox]::new("Account Locked")
      $state.chkSearchLocked.X=2
      $state.chkSearchLocked.Y=[Terminal.Gui.Pos]::Bottom($state.txtSearchOutput)+1
      $state.chkSearchLocked.CanFocus=$true
      $view.Add($state.chkSearchLocked)

      $state.chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Account Disabled")
      $state.chkSearchDisabled.X=2
      $state.chkSearchDisabled.Y=[Terminal.Gui.Pos]::Bottom($state.chkSearchLocked)+1
      $state.chkSearchDisabled.CanFocus=$true
      $view.Add($state.chkSearchDisabled)

      $btnDoSearch = [Terminal.Gui.Button]::new("Search")
      $btnDoSearch.X=48
      $btnDoSearch.Y=3
      $view.Add($btnDoSearch)

      ## Auto-populate search results
      [Terminal.Gui.Application]::MainLoop.Invoke({
        if ($user) {
          $lines = @()
          $user.PSObject.Properties | ForEach-Object {
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
          $state.chkSearchLocked.Checked = [bool]($user.Locked)
          $state.chkSearchDisabled.Checked = [bool]($user.Disabled)
        }
      }.GetNewClosure())
    }
  }

  ## ==================== Apply Logic ====================
  $applyLogic = {
    param($user, $state)

    Debug-Log ": Apply clicked - saving changes" -Type "Info"

    try {
      $changesMade = $false

      # Check SamAccountName change
      if ($state.txtSamAccountName) {
        $newSamAccountName = $state.txtSamAccountName.Text.ToString().Trim()
        if ($newSamAccountName -ne $user.SamAccountName -and -not [string]::IsNullOrWhiteSpace($newSamAccountName)) {
          if ($Script:DemoMode) {
            $user.SamAccountName = $newSamAccountName
            Debug-Log ": Updated SamAccountName to '$newSamAccountName' (demo)" -Type "Success"
            $changesMade = $true
          } else {
            Set-ADUser -Identity $user.SamAccountName -SamAccountName $newSamAccountName -ErrorAction Stop
            $user.SamAccountName = $newSamAccountName
            Debug-Log ": Updated SamAccountName to '$newSamAccountName'" -Type "Success"
            $changesMade = $true
          }
        }
      }

      # Check UserPrincipalName change
      if ($state.txtUserPrincipalName) {
        $newUPN = $state.txtUserPrincipalName.Text.ToString().Trim()
        if ($newUPN -ne $user.UserPrincipalName -and -not [string]::IsNullOrWhiteSpace($newUPN)) {
          if ($Script:DemoMode) {
            $user.UserPrincipalName = $newUPN
            Debug-Log ": Updated UPN to '$newUPN' (demo)" -Type "Success"
            $changesMade = $true
          } else {
            Set-ADUser -Identity $user.SamAccountName -UserPrincipalName $newUPN -ErrorAction Stop
            $user.UserPrincipalName = $newUPN
            Debug-Log ": Updated UPN to '$newUPN'" -Type "Success"
            $changesMade = $true
          }
        }
      }

      # Check DisplayName change
      if ($state.txtDisplayName) {
        $newDisplayName = $state.txtDisplayName.Text.ToString().Trim()
        if ($newDisplayName -ne $user.DisplayName -and -not [string]::IsNullOrWhiteSpace($newDisplayName)) {
          if ($Script:DemoMode) {
            $user.DisplayName = $newDisplayName
            $changesMade = $true
          } else {
            Set-ADUser -Identity $user.SamAccountName -DisplayName $newDisplayName -ErrorAction Stop
            $user.DisplayName = $newDisplayName
            $changesMade = $true
          }
        }
      }

      # Check Email change
      if ($state.txtEmail) {
        $newEmail = $state.txtEmail.Text.ToString().Trim()
        $currentEmail = if ($user.EmailAddress) { $user.EmailAddress } elseif ($user.mail) { $user.mail } else { "" }
        if ($newEmail -ne $currentEmail) {
          if ($Script:DemoMode) {
            $user.EmailAddress = $newEmail
            $user.mail = $newEmail
            $changesMade = $true
          } else {
            Set-ADUser -Identity $user.SamAccountName -EmailAddress $newEmail -ErrorAction Stop
            $user.EmailAddress = $newEmail
            $changesMade = $true
          }
        }
      }

      if ($changesMade) {
        Show-Modal "Success" "Changes applied successfully"
      } else {
        Show-Modal "Info" "No changes to apply"
      }

    } catch {
      Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
      Debug-Log ": Apply failed: $($_.Exception.Message)" -Type "Error"
    }
  }

  ## ==================== Create and Show Dialog ====================
  $tabs = @($generalTab, $accountTab, $addressTab, $profileTab, $organizationTab, $memberOfTab, $searchTab)

  New-PropertiesDialog -Title "User Properties - $($user.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $user -OnApply $applyLogic
}

## ==================== AUDIT LOG VIEWER ====================

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
        foreach ($char in $objectName.ToCharArray()) {
            $hash += [int]$char
        }

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

            $events = Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue |
                      Where-Object { $_.Message -match [regex]::Escape($objectName) } |
                      Select-Object -First 50

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
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnClose)

    ## Run dialog
    [Terminal.Gui.Application]::Run($dlg)

    Debug-Log ": Audit log dialog closed" -Type "Info"
}


## ==================== BUTTON ADDITION TO PROPERTIES DIALOGS ====================
##
## ADD THIS BUTTON TO THE GENERAL TAB OF:
## - Show-UserPropertiesDialog
## - Show-GroupPropertiesDialog
## - Show-ComputerPropertiesDialog
##
## Insert near the bottom of the General tab builder, after other fields:

## View Audit Log button
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
        if ($obj.PSObject.Properties.Match('SamAccountName') -and
            -not $obj.PSObject.Properties.Match('ComputerType')) {
            $hasUsers = $true
        }
        elseif ($obj.PSObject.Properties.Match('Members')) {
            $hasGroups = $true
        }
        elseif ($obj.PSObject.Properties.Match('ComputerType')) {
            $hasComputers = $true
        }
    }

    ## Define available attributes by object type
    $userAttributes = @(
        'DisplayName', 'Description', 'EmailAddress', 'Title', 'Department',
        'Company', 'Manager', 'OfficePhone', 'MobilePhone', 'StreetAddress',
        'City', 'PostalCode', 'Country',
        '--- PASSWORD OPTIONS ---',
        'ChangePasswordAtLogon', 'ResetPassword'
    )

    $groupAttributes = @(
        'Description', 'Email', 'ManagedBy'
    )

    $computerAttributes = @(
        'Description', 'Location'
    )

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
        $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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

        ## Validate attribute is valid for this specific object
        $validForObject = $false
        if ($objectType -eq 'User' -and $Attribute -in $userAttributes) { $validForObject = $true }
        if ($objectType -eq 'Group' -and $Attribute -in $groupAttributes) { $validForObject = $true }
        if ($objectType -eq 'Computer' -and $Attribute -in $computerAttributes) { $validForObject = $true }

        if (-not $validForObject) {
            $failCount++
            $errors += "${name}: Attribute '$Attribute' not valid for $objectType"
            Debug-Log ":   Skipped $name - attribute not valid for $objectType" -Type "Warn"
            continue
        }

        try {
            if ($Script:DemoMode) {
                ## Demo mode - update object property
                if ($Attribute -eq 'ChangePasswordAtLogon') {
                    $boolValue = [bool]::Parse($Value)
                    if ($obj.PSObject.Properties.Match('ChangePasswordAtLogon')) {
                        $obj.ChangePasswordAtLogon = $boolValue
                    } else {
                        $obj | Add-Member -NotePropertyName 'ChangePasswordAtLogon' -NotePropertyValue $boolValue -Force
                    }
                    $successCount++
                    Debug-Log ":   Updated $objectType '$name': ChangePasswordAtLogon = $boolValue (demo mode)" -Type "Info"
                }
                elseif ($Attribute -eq 'ResetPassword') {
                    $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
                        ## Generate random password
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
                }
                else {
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
                ## Production mode - use AD cmdlets
                if ($Attribute -eq 'ChangePasswordAtLogon') {
                    $boolValue = [bool]::Parse($Value)
                    Set-ADUser -Identity $obj.SamAccountName -ChangePasswordAtLogon $boolValue -ErrorAction Stop
                    $successCount++
                    Debug-Log ":   Updated $objectType '$name': ChangePasswordAtLogon = $boolValue in AD" -Type "Success"
                }
                elseif ($Attribute -eq 'ResetPassword') {
                    $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
                        ## Generate random password
                        $pw = -join ((65..90) + (97..122) + (48..57) + (33,35,36,37,38,42,63) | Get-Random -Count 12 | ForEach-Object {[char]$_})
                        ConvertTo-SecureString -String $pw -AsPlainText -Force
                    } else {
                        ConvertTo-SecureString -String $Value -AsPlainText -Force
                    }
                    Set-ADAccountPassword -Identity $obj.SamAccountName -NewPassword $passwordValue -Reset -ErrorAction Stop
                    $successCount++
                    Debug-Log ":   Reset password for $objectType '$name' in AD" -Type "Success"
                }
                else {
                    ## Standard attribute
                    $params = @{
                        Identity = if ($objectType -eq 'Group') { $obj.Name } else { $obj.SamAccountName }
                        ErrorAction = 'Stop'
                    }

                    ## Map attribute names to AD cmdlet parameters
                    switch ($Attribute) {
                        'EmailAddress' { $params['EmailAddress'] = $Value }
                        'Email' { $params['EmailAddress'] = $Value }
                        'DisplayName' { $params['DisplayName'] = $Value }
                        'Description' { $params['Description'] = $Value }
                        'Title' { $params['Title'] = $Value }
                        'Department' { $params['Department'] = $Value }
                        'Company' { $params['Company'] = $Value }
                        'Manager' { $params['Manager'] = $Value }
                        'ManagedBy' { $params['ManagedBy'] = $Value }
                        'OfficePhone' { $params['OfficePhone'] = $Value }
                        'MobilePhone' { $params['MobilePhone'] = $Value }
                        'StreetAddress' { $params['StreetAddress'] = $Value }
                        'City' { $params['City'] = $Value }
                        'PostalCode' { $params['PostalCode'] = $Value }
                        'Country' { $params['Country'] = $Value }
                        'Location' { $params['Location'] = $Value }
                    }

                    ## Execute appropriate cmdlet
                    switch ($objectType) {
                        'User' { Set-ADUser @params }
                        'Group' { Set-ADGroup @params }
                        'Computer' { Set-ADComputer @params }
                    }

                    $successCount++
                    Debug-Log ":   Updated $objectType '$name': $Attribute = '$Value' in AD" -Type "Success"
                }
            }

        } catch {
            $failCount++
            $errors += "${name}: $($_.Exception.Message)"
            Debug-Log ":   Failed to update ${name}: $($_.Exception.Message)" -Type "Error"
        }
    }

    ## Show results
    $msg = "Successfully updated $successCount object(s)"
    if ($failCount -gt 0) {
        $msg += "`n`nFailed: $failCount"
        if ($errors.Count -gt 0 -and $errors.Count -le 5) {
            $msg += "`n`nErrors:`n" + ($errors -join "`n")
        }
    }

    Show-Modal $(if ($failCount -eq 0) { "Success" } else { "Update Complete" }) $msg

    ## Refresh UI
    Refresh-Data -domain $Script:CurrentDomain
    Build-Tree -domain $Script:CurrentDomain
    if ($Script:FilterStatusLabel) {
        Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
    }

    Debug-Log ": Bulk attribute update completed - $successCount succeeded, $failCount failed" -Type "Success"
}

## ------------------------{ Apply Object Changes }------------------------
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
            # Production mode - use AD cmdlets
            switch ($ObjectType) {
                'User' {
                    $setParams = @{ Identity = $Object.SamAccountName }

                    if ($State.txtDisplayName) { $setParams['DisplayName'] = $State.txtDisplayName.Text.ToString() }
                    if ($State.txtDescription) { $setParams['Description'] = $State.txtDescription.Text.ToString() }
                    if ($State.txtEmail) { $setParams['EmailAddress'] = $State.txtEmail.Text.ToString() }
                    if ($State.txtOfficePhone) { $setParams['OfficePhone'] = $State.txtOfficePhone.Text.ToString() }
                    if ($State.txtMobilePhone) { $setParams['MobilePhone'] = $State.txtMobilePhone.Text.ToString() }
                    if ($State.txtStreet) { $setParams['StreetAddress'] = $State.txtStreet.Text.ToString() }
                    if ($State.txtCity) { $setParams['City'] = $State.txtCity.Text.ToString() }
                    if ($State.txtPostal) { $setParams['PostalCode'] = $State.txtPostal.Text.ToString() }
                    if ($State.txtCountry) { $setParams['Country'] = $State.txtCountry.Text.ToString() }
                    if ($State.txtTitle) { $setParams['Title'] = $State.txtTitle.Text.ToString() }
                    if ($State.txtDept) { $setParams['Department'] = $State.txtDept.Text.ToString() }
                    if ($State.txtCompany) { $setParams['Company'] = $State.txtCompany.Text.ToString() }
                    if ($State.txtManager) { $setParams['Manager'] = $State.txtManager.Text.ToString() }

                    Set-ADUser @setParams -ErrorAction Stop

                    # Handle SamAccountName/UPN changes separately
                    if ($State.txtSamAccountName) {
                        $newSam = $State.txtSamAccountName.Text.ToString().Trim()
                        if ($newSam -ne $Object.SamAccountName) {
                            Set-ADUser -Identity $Object.SamAccountName -SamAccountName $newSam -ErrorAction Stop
                            $Object.SamAccountName = $newSam
                        }
                    }

                    if ($State.txtUserPrincipalName) {
                        $newUPN = $State.txtUserPrincipalName.Text.ToString().Trim()
                        if ($newUPN -ne $Object.UserPrincipalName) {
                            Set-ADUser -Identity $Object.SamAccountName -UserPrincipalName $newUPN -ErrorAction Stop
                            $Object.UserPrincipalName = $newUPN
                        }
                    }

                    # Handle account status
                    if ($State.chkEnabled) {
                        if ($State.chkEnabled.Checked -and -not $Object.Enabled) {
                            Enable-ADAccount -Identity $Object.SamAccountName -ErrorAction Stop
                        } elseif (-not $State.chkEnabled.Checked -and $Object.Enabled) {
                            Disable-ADAccount -Identity $Object.SamAccountName -ErrorAction Stop
                        }
                    }
                }

                'Group' {
                    $setParams = @{ Identity = $Object.Name }

                    if ($State.txtDescription) { $setParams['Description'] = $State.txtDescription.Text.ToString() }
                    if ($State.txtEmail -and $State.txtEmail.Text.ToString()) {
                        $setParams['Replace'] = @{ mail = $State.txtEmail.Text.ToString() }
                    }
                    if ($State.txtManagedBy -and $State.txtManagedBy.Text.ToString()) {
                        $setParams['ManagedBy'] = $State.txtManagedBy.Text.ToString()
                    }

                    Set-ADGroup @setParams -ErrorAction Stop
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
                        if ($adOU -and $newDesc) {
                            Set-ADOrganizationalUnit -Identity $adOU.DistinguishedName -Description $newDesc -ErrorAction Stop
                        }
                    }

                    $Object.Description = $newDesc
                }

                'Computer' {
                    $setParams = @{ Identity = $Object.SamAccountName }

                    if ($State.txtDescription) { $setParams['Description'] = $State.txtDescription.Text.ToString() }
                    if ($State.txtLocation) { $setParams['Location'] = $State.txtLocation.Text.ToString() }

                    Set-ADComputer @setParams -ErrorAction Stop
                }
            }

            Debug-Log "SUCCESS: $ObjectType changes applied to AD" -Type "Success"
            Show-Modal "Success" "Changes applied successfully"
        }

        # Refresh data
        Refresh-Data -domain $Script:CurrentDomain

        # Clear change tracking flags
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

## -----------------------{ Refresh Domain Data }----------------------
function Refresh-Data {
  param([string]$domain, [switch]$RebuildTree)

  ## If in demo mode, just reload demo data
  if ($Script:DemoMode) {
    try {
      Set-StatusBar "Refreshing demo data..." -spinner
      $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $domain

      if ($RebuildTree) {
        Set-StatusBar "Rebuilding tree..." -spinner

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

      Set-StatusBar "Demo data refreshed" -final
      return $true

    } catch {
      Debug-Log ("Demo refresh error: $_") -Type "Warn"
      Set-StatusBar "Refresh error" -final
      return $false
    }
  }

  ## Production mode - query AD
  try {
    Set-StatusBar "Loading DCs..." -spinner
    $dcs = Invoke-AD { Get-ADDomainController -Filter * -DomainName $domain -ErrorAction Stop } -SuppressError

    Set-StatusBar "Loading Users..." -spinner
    $users = Invoke-AD { Get-ADUser -Filter * -Server $domain -Properties DisplayName,EmailAddress,Title,Department,Enabled,LockedOut,DistinguishedName -ErrorAction Stop } -SuppressError

    Set-StatusBar "Loading Groups..." -spinner
    $groups = Invoke-AD { Get-ADGroup -Filter * -Server $domain -Properties Description,GroupCategory,GroupScope,Members,DistinguishedName -ErrorAction Stop } -SuppressError

    Set-StatusBar "Loading Computers..." -spinner
    $computers = Invoke-AD { Get-ADComputer -Filter * -Server $domain -Properties OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,DistinguishedName -ErrorAction Stop } -SuppressError

    ## Check results
    if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
      Debug-Log ("Refresh failed - one or more queries returned null") -Type "Warn"
      Set-StatusBar "Refresh failed - check logs" -final
      return $false
    }

    Set-StatusBar "Converting data..." -spinner
    $converted = Convert-DataToADObjects -Users $users -DCs $dcs -Groups $groups -Computers $computers -Domain $domain

    if ($RebuildTree) {
      Set-StatusBar "Rebuilding tree..." -spinner

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

  Set-StatusBar "Refresh complete" -final
  return $true

  } catch {
    Debug-Log ("Refresh error: $_") -Type "Warn"
    Set-StatusBar "Refresh error" -final
    return $false
  }
}

## Is it a special day...?
## Determines which emoji to use for the Active Directory window title. Uses Unicode escapes for flags to avoid editor issues.
function Initialize-DirectoryEmoji {
  param(
    [DateTime]$Date = (Get-Date)
  )

  $month = $Date.Month
  $day   = $Date.Day

  ## Default: card index
  $emoji = "🗂️"

  switch ($true) {
    { $month -eq 4  -and $day -eq 9  } { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }      ## 9th Apr Danmarks besættelse (liberation day)
    { $month -eq 5  -and $day -eq 4  } { $emoji = "🕯️" ; break }                     ## 4th May Candle for Besættelsen
    { $month -eq 5  -and $day -eq 5  } { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }      ## 5th May Constitution day in Denmark
    { $month -eq 6  -and $day -eq 21 } { $emoji = "`u{1F1EC}`u{1F1F1}" ; break }     ## 21st May Grønland Day
    { $month -eq 7  -and $day -eq 29 } { $emoji = "`u{1F1EB}`u{1F1F4}" ; break }     ## 29th Jul Faroe Islands
    { $month -eq 9  -and $day -eq 9  } { $emoji = "`u{1F1E9}`u{1F1EA}" ; break }      ## 9th Nov Erich Honecker leck mich am Arsch!
    { $month -eq 12 -and ($day -eq 24 -or $day -eq 25) } { $emoji = "🎄" ; break }  ## 24th/25th Dec Tree for Jul / Christmas
  }

  Debug-Log ("Today's emoji is: $emoji") -Type "info"
  $Script:DirectoryEmoji = $emoji
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
      ## ========== Eurythmics – UK/Scotland/Aberdeen (Satellite Office) ==========
      @{ Name = 'Annie Lennox'             ; SamAccountName = 'annie.lennox'      ; UserPrincipalName = 'annie.lennox@example.org'       ; OU = @('Locations','UK','Scotland','Aberdeen','Eurythmics')                           ; Groups = @('Eurythmics','Vocalists','Keyboardists')   ; Title = 'Lead Vocalist/Keyboardist'  ; Email = 'annie.lennox@example.org'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aberdeen Office'   ; Phone = '+44 1224 496 010' ; MobilePhone = '+44 7700 941001' ; Street = '210 Union Street'         ; City = 'Aberdeen'    ; PostalCode = 'AB10 1TL'    ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist and co-founder of Eurythmics'                   },
      @{ Name = 'Dave Stewart'             ; SamAccountName = 'dave.stewart'      ; UserPrincipalName = 'dave.stewart@example.org'       ; OU = @('Locations','UK','Scotland','Aberdeen','Eurythmics')                           ; Groups = @('Eurythmics','Guitarists','Producers')     ; Title = 'Guitarist and Producer'     ; Email = 'dave.stewart@example.org'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aberdeen Office'   ; Phone = '+44 1224 496 011' ; MobilePhone = '+44 7700 941002' ; Street = '18 Rosemount Place'       ; City = 'Aberdeen'    ; PostalCode = 'AB25 2XP'    ; Company = 'Example Music Ltd'   ; Manager = 'Annie Lennox'     ; Description = 'Guitarist, songwriter and producer for Eurythmics'            },

      ## ========== Deacon Blue – UK/Scotland/Dundee ==========
      @{ Name = 'Ricky Ross'               ; SamAccountName = 'ricky.ross'        ; UserPrincipalName = 'ricky.ross@example.net'         ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Lead Vocalist'               ; Email = 'ricky.ross@example.net'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 001' ; MobilePhone = '+44 7700 392001' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist for Deacon Blue'                                },
      @{ Name = 'Lorraine McIntosh'        ; SamAccountName = 'lorraine.mcintosh' ; UserPrincipalName = 'lorraine.mcintosh@example.net'  ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Vocalist'                    ; Email = 'lorraine.mcintosh@example.net'    ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 002' ; MobilePhone = '+44 7700 392002' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = 'Ricky Ross'       ; Description = 'Vocalist for Deacon Blue'                                     },
      @{ Name = 'Dougie Vipond'            ; SamAccountName = 'dougie.vipond'     ; UserPrincipalName = 'dougie.vipond@example.net'      ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Drummer'                     ; Email = 'dougie.vipond@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 003' ; MobilePhone = '+44 7700 392003' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = 'Ricky Ross'       ; Description = 'Drummer for Deacon Blue'                                      },
      @{ Name = 'James Prime'              ; SamAccountName = 'james.prime'       ; UserPrincipalName = 'james.prime@example.net'        ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Keyboardist'                 ; Email = 'james.prime@example.net'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 004' ; MobilePhone = '+44 7700 392004' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = 'Ricky Ross'       ; Description = 'Keyboardist for Deacon Blue'                                  },
      @{ Name = 'Ewen Vernal'              ; SamAccountName = 'ewen.vernal'       ; UserPrincipalName = 'ewen.vernal@example.net'        ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Bassist'                     ; Email = 'ewen.vernal@example.net'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 005' ; MobilePhone = '+44 7700 392005' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = 'Ricky Ross'       ; Description = 'Bassist for Deacon Blue'                                      },
      @{ Name = 'Graeme Kelling'           ; SamAccountName = 'graeme.kelling'    ; UserPrincipalName = 'graeme.kelling@example.net'     ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Guitarist'                   ; Email = 'graeme.kelling@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 006' ; MobilePhone = '+44 7700 392006' ; Street = '1 Tannadice Street'       ; City = 'Dundee'      ; PostalCode = 'DD3 7JW'     ; Company = 'Example Music Ltd'   ; Manager = 'Ricky Ross'       ; Description = 'Guitarist for Deacon Blue'                                    },

      ## ========== Marillion (UK/Scotland/Edinburgh) ==========
      @{ Name = 'Derek Dick'               ; SamAccountName = 'fish'              ; UserPrincipalName = 'fish@example.com'               ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Vocalists')                   ; Title = 'Lead Vocalist'               ; Email = 'fish@example.com'                 ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0221' ; MobilePhone = '+44 7700 222221' ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Former lead vocalist (Fish) for Marillion (1981-1988)'        },
      @{ Name = 'Steve Rothery'            ; SamAccountName = 'steve.rothery'     ; UserPrincipalName = 'steve.rothery@example.com'      ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Guitarists')                  ; Title = 'Lead Guitarist'              ; Email = 'steve.rothery@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0222' ; MobilePhone = '+44 7700 222222' ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music Ltd'   ; Manager = 'Derek Dick'       ; Description = 'Lead guitarist and founding member of Marillion'              },
      @{ Name = 'Pete Trewavas'            ; SamAccountName = 'pete.trewavas'     ; UserPrincipalName = 'pete.trewavas@example.com'      ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Guitarists')                  ; Title = 'Bassist'                     ; Email = 'pete.trewavas@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0223' ; MobilePhone = '+44 7700 222223' ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music Ltd'   ; Manager = 'Derek Dick'       ; Description = 'Bassist and founding member of Marillion'                     },
      @{ Name = 'Mark Kelly'               ; SamAccountName = 'mark.kelly'        ; UserPrincipalName = 'mark.kelly@example.com'         ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Keyboards')                   ; Title = 'Keyboardist'                 ; Email = 'mark.kelly@example.com'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0224' ; MobilePhone = '+44 7700 222224' ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music Ltd'   ; Manager = 'Derek Dick'       ; Description = 'Keyboardist and founding member of Marillion'                 },
      @{ Name = 'Ian Mosley'               ; SamAccountName = 'ian.mosley'        ; UserPrincipalName = 'ian.mosley@example.com'         ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Percussion')                  ; Title = 'Drummer'                     ; Email = 'ian.mosley@example.com'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0225' ; MobilePhone = '+44 7700 222225' ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music Ltd'   ; Manager = 'Derek Dick'       ; Description = 'Drummer for Marillion (joined 1984)'                          },

      ## ========== Proclaimers (UK/Scotland/Edinburgh) ==========
      @{ Name = 'Craig Reid'               ; SamAccountName = 'craig.reid'        ; UserPrincipalName = 'craig.reid@example.net'         ; OU = @('Locations','CA','Ontario','Brockville','The Proclaimers')                     ; Groups = @('The Proclaimers','VPN-Users')             ; Title = 'Vocalist / Guitarist'        ; Email = 'craig.reid@example.net'           ; Country='UK'   ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0101' ; MobilePhone = '+44 7700 496011' ; Street = '12 Albion Place'          ; City = 'Edinburgh'   ; PostalCode = 'EH7 5DG'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description='Founding member of The Proclaimers'                             },
      @{ Name = 'Charlie Reid'             ; SamAccountName = 'charlie.reid'      ; UserPrincipalName = 'charlie.reid@example.net'       ; OU = @('Locations','US','Florida','Miami','The Proclaimers')                          ; Groups = @('The Proclaimers','VPN-Users')             ; Title = 'Vocalist / Guitarist'        ; Email = 'charlie.reid@example.net'         ; Country='UK'   ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0102' ; MobilePhone = '+44 7700 496012' ; Street = '12 Albion Place'          ; City = 'Edinburgh'   ; PostalCode = 'EH7 5DG'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description='Founding member of The Proclaimers'                             },

      ## ========== Ultravox – UK/Scotland/Edinburgh (And Vienna) ==========
      @{ Name = 'Midge Ure'                ; SamAccountName = 'midge.ure'         ; UserPrincipalName = 'midge.ure@example.org'          ; OU = @('Locations','UK','Scotland','Edinburgh','Ultravox')                            ; Groups=@('Ultravox','Sales')                          ; Title = 'Vocalist / Guitarist'        ; Email = 'midge.ure@example.org'            ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0226' ; MobilePhone = ''                ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music UK'    ; Manager = ''                 ; Description = 'Vocalist and guitarist for Ultravox'                          },
      @{ Name = 'Billy Currie'             ; SamAccountName = 'billy.currie'      ; UserPrincipalName = 'billy.currie@example.org'       ; OU = @('Locations','UK','Scotland','Edinburgh','Ultravox')                            ; Groups=@('Ultravox','Sales')                          ; Title = 'Keyboardist / Violinist'     ; Email = 'billy.currie@example.org'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0227' ; MobilePhone = ''                ; Street = '22 Tynecastle Street'     ; City = 'Edinburgh'   ; PostalCode = 'EH1 2BB'     ; Company = 'Example Music UK'    ; Manager = 'Midge Ure'        ; Description = 'Keyboardist and violinist for Ultravox'                       },
      @{ Name = 'Chris Cross'              ; SamAccountName = 'chris.cross'       ; UserPrincipalName = 'chris.cross@example.org'        ; OU = @('Locations','Austria','Vienna','Ultravox')                                     ; Groups=@('Ultravox','Sales')                          ; Title = 'Bassist'                     ; Email = 'chris.cross@example.org'          ; Country = 'AT' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Vienna Office'     ; Phone = '+44 131 496 0228' ; MobilePhone = ''                ; Street = 'Am Hauptbahnhof'          ; City = 'Vienna'      ; PostalCode = '1100 Wien'   ; Company = 'Example Music AT'    ; Manager = 'Midge Ure'        ; Description = 'Bassist for Ultravox'                                         },
      @{ Name = 'Warren Cann'              ; SamAccountName = 'warren.cann'       ; UserPrincipalName = 'warren.cann@example.org'        ; OU = @('Locations','Austria','Vienna','Ultravox')                                     ; Groups=@('Ultravox','Sales')                          ; Title = 'Drummer'                     ; Email = 'warren.cann@example.org'          ; Country = 'AT' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Vienna Office'     ; Phone = '+44 131 496 0229' ; MobilePhone = ''                ; Street = 'Am Hauptbahnhof'          ; City = 'Vienna'      ; PostalCode = '1100 Wien'   ; Company = 'Example Music AT'    ; Manager = 'Midge Ure'        ; Description = 'Drummer for Ultravox'                                         },

      ## ========== Altered Images – UK/Scotland/Glasgow ==========
      @{ Name = 'Clare Grogan'             ; SamAccountName = 'clare.grogan'      ; UserPrincipalName = 'clare.grogan@example.org'       ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Vocalists')              ; Title = 'Lead Vocalist'               ; Email = 'clare.grogan@example.org'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0101' ; MobilePhone = '+44 7700 931001' ; Street = '150 Sauchiehall Street'   ; City = 'Glasgow'     ; PostalCode = 'G2 3EL'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist of Altered Images'                              },
      @{ Name = 'Johnny McElhone'          ; SamAccountName = 'johnny.mcelhone'   ; UserPrincipalName = 'johnny.mcelhone@example.org'    ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Bassists')               ; Title = 'Bassist'                     ; Email = 'johnny.mcelhone@example.org'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0102' ; MobilePhone = '+44 7700 931002' ; Street = '42 Hope Street'           ; City = 'Glasgow'     ; PostalCode = 'G2 6AE'      ; Company = 'Example Music Ltd'   ; Manager = 'Clare Grogan'     ; Description = 'Bassist and songwriter for Altered Images'                    },
      @{ Name = 'Jim McKinven'             ; SamAccountName = 'jim.mckinven'      ; UserPrincipalName = 'jim.mckinven@example.org'       ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Guitarists')             ; Title = 'Guitarist'                   ; Email = 'jim.mckinven@example.org'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0103' ; MobilePhone = '+44 7700 931003' ; Street = '77 Bath Street'           ; City = 'Glasgow'     ; PostalCode = 'G2 2EN'      ; Company = 'Example Music Ltd'   ; Manager = 'Clare Grogan'     ; Description = 'Guitarist for Altered Images'                                 },
      @{ Name = 'Michael Anderson'         ; SamAccountName = 'michael.anderson'  ; UserPrincipalName = 'michael.anderson@example.org'   ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Drummers')               ; Title = 'Drummer'                     ; Email = 'michael.anderson@example.org'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0104' ; MobilePhone = '+44 7700 931004' ; Street = '305 Argyle Street'        ; City = 'Glasgow'     ; PostalCode = 'G2 8DL'      ; Company = 'Example Music Ltd'   ; Manager = 'Clare Grogan'     ; Description = 'Drummer for Altered Images (aka Tich Anderson)'               },

      ## ========== Simple Minds (UK/Scotland/Glasgow) ==========
      @{ Name = 'Jim Kerr'                 ; SamAccountName = 'jkerr'             ; UserPrincipalName = 'jkerr@example.com'              ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Vocalists')                ; Title = 'Lead Vocalist'               ; Email = 'jim.kerr@example.com'             ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0111' ; MobilePhone = '+44 7700 111111' ; Street = '1 Sauchiehall Street'     ; City = 'Glasgow'     ; PostalCode = 'G1 1AA'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist for Simple Minds'                               },
      @{ Name = 'Charlie Burchill'         ; SamAccountName = 'charlie.b'         ; UserPrincipalName = 'charlie.b@example.com'          ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Guitarists')               ; Title = 'Lead Guitarist'              ; Email = 'charlie.b@example.com'            ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0112' ; MobilePhone = '+44 7700 111112' ; Street = '1 Sauchiehall Street'     ; City = 'Glasgow'     ; PostalCode = 'G1 1AA'      ; Company = 'Example Music Ltd'   ; Manager = 'Jim Kerr'         ; Description = 'Guitarist and founding member of Simple Minds'                },
      @{ Name = 'Mel Gaynor'               ; SamAccountName = 'mel.gaynor'        ; UserPrincipalName = 'mel.gaynor@example.com'         ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Percussion')               ; Title = 'Drummer'                     ; Email = 'mel.gaynor@example.com'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0113' ; MobilePhone = '+44 7700 111113' ; Street = '1 Sauchiehall Street'     ; City = 'Glasgow'     ; PostalCode = 'G1 1AA'      ; Company = 'Example Music Ltd'   ; Manager = 'Jim Kerr'         ; Description = 'Drummer for Simple Minds'                                     },
      @{ Name = 'Mick MacNeil'             ; SamAccountName = 'mick.macneil'      ; UserPrincipalName = 'mick.macneil@example.com'       ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Musicians','Former Staff') ; Title = 'Keyboardist (Former)'        ; Email = 'mick.macneil@example.com'         ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0120' ; MobilePhone = '+44 7700 111120' ; Street = '1 Sauchiehall Street'     ; City = 'Glasgow'     ; PostalCode = 'G1 1AA'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Former keyboardist for Simple Minds (1977-1990)'              },
      @{ Name = 'Derek Forbes'             ; SamAccountName = 'derek.forbes'      ; UserPrincipalName = 'derek.forbes@example.com'       ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Musicians','Former Staff') ; Title = 'Bassist (Former)'            ; Email = 'derek.forbes@example.com'         ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0121' ; MobilePhone = '+44 7700 111121' ; Street = '1 Sauchiehall Street'     ; City = 'Glasgow'     ; PostalCode = 'G1 1AA'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Former bassist for Simple Minds (1977-1985)'                  },

      ## ========== Wet Wet Wet — UK / Scotland / Clydebank ==========
      @{ Name = 'Marti Pellow'             ; SamAccountName = 'marti.pellow'      ; UserPrincipalName = 'marti.pellow@example.net'       ; OU=@('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                           ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Lead Vocalist'               ; Email = 'marti.pellow@example.net'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0201' ; MobilePhone = '+44 7700 496201' ; Street = 'Dumbarton Road 200'       ; City = 'Clydebank'   ; PostalCode = 'G81 1UE'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist of Wet Wet Wet'                                 },
      @{ Name = 'Graeme Clark'             ; SamAccountName = 'graeme.clark'      ; UserPrincipalName = 'graeme.clark@example.net'       ; OU=@('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                           ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Bassist'                     ; Email = 'graeme.clark@example.net'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0202' ; MobilePhone = '+44 7700 496202' ; Street = 'Dumbarton Road 201'       ; City = 'Clydebank'   ; PostalCode = 'G81 1UE'     ; Company = 'Example Music Ltd'   ; Manager = 'Marti Pellow'     ; Description = 'Bassist for Wet Wet Wet'                                      },
      @{ Name = 'Tommy Cunningham'         ; SamAccountName = 'tommy.cunningham'  ; UserPrincipalName = 'tommy.cunningham@example.net'   ; OU=@('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                           ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Drummer'                     ; Email = 'tommy.cunningham@example.net'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0203' ; MobilePhone = '+44 7700 496203' ; Street = 'Dumbarton Road 202'       ; City = 'Clydebank'   ; PostalCode = 'G81 1UE'     ; Company = 'Example Music Ltd'   ; Manager = 'Marti Pellow'     ; Description = 'Drummer for Wet Wet Wet'                                      },
      @{ Name = 'Neil Mitchell'            ; SamAccountName = 'neil.mitchell'     ; UserPrincipalName = 'neil.mitchell@example.net'      ; OU=@('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                           ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Keyboardist'                 ; Email = 'neil.mitchell@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0204' ; MobilePhone = '+44 7700 496204' ; Street = 'Dumbarton Road 203'       ; City = 'Clydebank'   ; PostalCode = 'G81 1UE'     ; Company = 'Example Music Ltd'   ; Manager = 'Marti Pellow'     ; Description = 'Keyboardist for Wet Wet Wet'                                  },

      ## ==========The Police - UK/England/Newcastle ==========
      @{ Name = 'Gordon Summer'            ; SamAccountName = 'sting'             ; UserPrincipalName = 'sting@example.org'              ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Vocalists','Bassists')        ; Title = 'Lead Vocalist & Bassist'     ; Email = 'sting@example.org'                ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0001' ; MobilePhone = '+44 7700 910001' ; Street = '14 Grey Street'           ; City = 'Newcastle'   ; PostalCode = 'NE1 6BH'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist and bassist of The Police'                      },
      @{ Name = 'Andy Summers'             ; SamAccountName = 'andy.summers'      ; UserPrincipalName = 'andy.summers@example.org'       ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Guitarists')                  ; Title = 'Guitarist'                   ; Email = 'andy.summers@example.org'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0002' ; MobilePhone = '+44 7700 910002' ; Street = '11 Dean Street'           ; City = 'Newcastle'   ; PostalCode = 'NE1 1PG'     ; Company = 'Example Music Ltd'   ; Manager = 'Sting'            ; Description = 'Guitarist for The Police'                                     },
      @{ Name = 'Stewart Copeland'         ; SamAccountName = 'stewart.copeland'  ; UserPrincipalName = 'stewart.copeland@example.org'   ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Drummers')                    ; Title = 'Drummer'                     ; Email = 'stewart.copeland@example.org'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0003' ; MobilePhone = '+44 7700 910003' ; Street = '5 Collingwood Street'     ; City = 'Newcastle'   ; PostalCode = 'NE1 1JF'     ; Company = 'Example Music Ltd'   ; Manager = 'Sting'            ; Description = 'Drummer for The Police'                                       },

      ## ========== New Order – UK/England/Manchester ==========
      @{ Name = 'Bernard Sumner'           ; SamAccountName = 'bernard.sumner'    ; UserPrincipalName = 'bernard.sumner@example.org'     ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Vocalists','Guitarists')       ; Title = 'Lead Vocalist & Guitarist'   ; Email = 'bernard.sumner@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0161' ; MobilePhone = '+44 7700 951161' ; Street = 'Annalade Road 10'         ; City = 'Manchester'  ; PostalCode = 'M16 9AB'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist and guitarist for New Order'                    },
      @{ Name = 'Stephen Morris'           ; SamAccountName = 'stephen.morris'    ; UserPrincipalName = 'stephen.morris@example.org'     ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Drummers')                     ; Title = 'Drummer'                     ; Email = 'stephen.morris@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0162' ; MobilePhone = '+44 7700 951162' ; Street = 'Cromwell Road 22'         ; City = 'Manchester'  ; PostalCode = 'M16 9AB'     ; Company = 'Example Music Ltd'   ; Manager = 'Bernard Sumner'   ; Description = 'Drummer for New Order'                                        },
      @{ Name = 'Peter Hook'               ; SamAccountName = 'peter.hook'        ; UserPrincipalName = 'peter.hook@example.org'         ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Bassists')                     ; Title = 'Bassist'                     ; Email = 'peter.hook@example.org'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0163' ; MobilePhone = '+44 7700 951163' ; Street = 'Blake Road 5'             ; City = 'Manchester'  ; PostalCode = 'M16 9AB'     ; Company = 'Example Music Ltd'   ; Manager = 'Bernard Sumner'   ; Description = 'Bassist and co-founder of New Order'                          },
      @{ Name = 'Gillian Gilbert'          ; SamAccountName = 'gillian.gilbert'   ; UserPrincipalName = 'gillian.gilbert@example.org'    ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Keyboardists','Guitarists')    ; Title = 'Keyboardist & Guitarist'     ; Email = 'gillian.gilbert@example.org'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0164' ; MobilePhone = '+44 7700 951164' ; Street = 'Chorlton Road 18'         ; City = 'Manchester'  ; PostalCode = 'M16 9AB'     ; Company = 'Example Music Ltd'   ; Manager = 'Bernard Sumner'   ; Description = 'Keyboardist and guitarist for New Order'                      },

      ## ========== Echo And The Bunnymen - UK/England/Liverpool ==========
      @{ Name = 'Ian McCulloch'            ; SamAccountName = 'ian.mcculloch'     ; UserPrincipalName = 'ian.mcculloch@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Vocalists')        ; Title = 'Lead Vocalist'               ; Email = 'ian.mcculloch@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0001' ; MobilePhone = '+44 7700 900001' ; Street = '12 Mathew Street'         ; City = 'Liverpool'   ; PostalCode = 'L1 4ED'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist of Echo and The Bunnymen'                       },
      @{ Name = 'Will Sergeant'            ; SamAccountName = 'will.sergeant'     ; UserPrincipalName = 'will.sergeant@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Guitarists')       ; Title = 'Guitarist'                   ; Email = 'will.sergeant@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0002' ; MobilePhone = '+44 7700 900002' ; Street = '22 Bold Street'           ; City = 'Liverpool'   ; PostalCode = 'L1 4HR'      ; Company = 'Example Music Ltd'   ; Manager = 'Ian McCulloch'    ; Description = 'Guitarist for Echo and The Bunnymen'                          },
      @{ Name = 'Les Pattinson'            ; SamAccountName = 'les.pattinson'     ; UserPrincipalName = 'les.pattinson@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Bassists')         ; Title = 'Bass Guitarist'              ; Email = 'les.pattinson@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 6003' ; MobilePhone = '+44 7700 900003' ; Street = '8 Seel Street'            ; City = 'Liverpool'   ; PostalCode = 'L1 4BE'      ; Company = 'Example Music Ltd'   ; Manager = 'Ian McCulloch'    ; Description = 'Bass guitarist for Echo and The Bunnymen'                     },
      @{ Name = 'Pete de Freitas'          ; SamAccountName = 'pete.defreitas'    ; UserPrincipalName = 'pete.defreitas@example.org'     ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Drummers')         ; Title = 'Drummer'                     ; Email = 'pete.defreitas@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0004' ; MobilePhone = '+44 7700 900004' ; Street = '5 Dale Street'            ; City = 'Liverpool'   ; PostalCode = 'L2 2EH'      ; Company = 'Example Music Ltd'   ; Manager = 'Ian McCulloch'    ; Description = 'Drummer for Echo and The Bunnymen'                            },

      ## ========== UB40 — UK / England / Birmingham ==========
      @{ Name = 'Ali Campbell'             ; SamAccountName = 'ali.campbell'      ; UserPrincipalName = 'ali.campbell@example.net'       ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Lead Vocalist'               ; Email = 'ali.campbell@example.net'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone = '+44 121 496 0101' ; MobilePhone = '+44 7700 921601' ; Street = '40 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead voca l ist of UB40'                                      },
      @{ Name = 'Robin Campbell'           ; SamAccountName = 'robin.campbell'    ; UserPrincipalName = 'robin.campbell@example.net'     ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Guitarist'                   ; Email = 'robin.campbell@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0102' ; MobilePhone = '+44 7700 921602' ; Street = '41 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Guitarist for UB40'                                           },
      @{ Name = 'Brian Travers'            ; SamAccountName = 'brian.travers'     ; UserPrincipalName = 'brian.travers@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Saxophonist'                 ; Email = 'brian.travers@example.net'        ; Country = 'UK' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0103' ; MobilePhone = '+44 7700 921603' ; Street = '42 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Saxophonist for UB40 (account disabled)'                      },
      @{ Name = 'Earl Falconer'            ; SamAccountName = 'earl.falconer'     ; UserPrincipalName = 'earl.falconer@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Bassist'                     ; Email = 'earl.falconer@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0104' ; MobilePhone = '+44 7700 921604' ; Street = '43 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Bassist for UB40'                                             },
      @{ Name = 'Norman Hassan'            ; SamAccountName = 'norman.hassan'     ; UserPrincipalName = 'norman.hassan@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Percussionist'               ; Email = 'norman.hassan@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0105' ; MobilePhone = '+44 7700 921605' ; Street = '44 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Percussionist for UB40'                                       },
      @{ Name = 'Terence Wilson'           ; SamAccountName = 'astro.wilson'      ; UserPrincipalName = 'astro@example.net'              ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Toaster / Trumpet'           ; Email = 'astro.wilson@example.net'         ; Country = 'UK' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0106' ; MobilePhone = '+44 7700 921606' ; Street = '45 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Toaster and trumpeter for UB40 (account disabled)'            },
      @{ Name = 'James Brown'              ; SamAccountName = 'james.brown'       ; UserPrincipalName = 'james.brown@example.net'        ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Drummer'                     ; Email = 'james.brown@example.net'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0107' ; MobilePhone = '+44 7700 921607' ; Street = '46 Broad Street'          ; City = 'Birmingham'  ; PostalCode = 'B1 2EU'      ; Company = 'Example Music Ltd'   ; Manager = 'Ali Campbell'     ; Description = 'Drummer for UB40'                                             },

      ## ========== Erasure (UK/England/London) ==========
      @{ Name = 'Andy Bell'                ; SamAccountName = 'andy.bell'         ; UserPrincipalName = 'andy.bell@example.com'          ; OU = @('Locations','UK','England','London','Erasure')                                 ; Groups = @('Erasure','Vocalists')                       ; Title = 'Lead Vocalist'               ; Email = 'andy.bell@example.com'            ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0011' ; MobilePhone = '+44 7700 333333' ; Street = '15 Carnaby Street'        ; City = 'London'      ; PostalCode = 'E1 7AA'      ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Lead vocalist for Erasure'                                    },
      @{ Name = 'Vince Clarke'             ; SamAccountName = 'vince.clarke'      ; UserPrincipalName = 'vince.clarke@example.com'       ; OU = @('Locations','UK','England','London','Erasure')                                 ; Groups = @('Erasure','Synth','Keyboards')               ; Title = 'Synth / Keyboardist'         ; Email = 'vince.clarke@example.com'         ; Country = 'UK' ; Disabled = $false ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0012' ; MobilePhone = '+44 7700 333334' ; Street = '15 Carnaby Street'        ; City = 'London'      ; PostalCode = 'E1 7AA'      ; Company = 'Example Music Ltd'   ; Manager = 'Andy Bell'        ; Description = 'Synthesizer pioneer - member of Depeche Mode & Erasure'       },

      ## ========== Mel And Kim – UK/England/London ==========
      @{ Name='Melanie Appleby'            ; SamAccountName='melanie.appleby'     ; UserPrincipalName='mel.appleby@example.com'          ; OU=@('Locations','UK','England','London','Sales')                                     ; Groups = @('Mel And Kim','Sales')                       ; Title = 'Singer and Dancer'           ; Email='mel.appleby@example.com'            ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Sales' ; Office = 'London Office'     ; Phone = '+44 207 496 0013' ; MobilePhone = '+44 7700 333335' ; Street = '15 Carnaby Street'        ; City ='London'       ; PostalCode = 'E1 7AA'      ; Company = 'Example Music Ltd'   ; Manager=''                   ; Description='Member of Mel and Kim, London sales team'                       },
      @{ Name='Kimberly Appleby'           ; SamAccountName='kim.appleby'         ; UserPrincipalName='kim.appleby@example.com'          ; OU=@('Locations','UK','England','London','Sales')                                     ; Groups = @('Mel And Kim','Sales')                       ; Title = 'Singer and Dancer'           ; Email='kim.appleby@example.com'            ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Sales' ; Office = 'London Office'     ; Phone = '+44 207 496 0014' ; MobilePhone = '+44 7700 333336' ; Street = '15 Carnaby Street'        ; City ='London'       ; PostalCode = 'E1 7AA'      ; Company = 'Example Music Ltd'   ; Manager=''                   ; Description='Member of Mel and Kim, London sales team'                       },

      ## ========== Depeche Mode (UK/England/London) ==========
      @{ Name = 'Dave Gahan'               ; SamAccountName = 'dave.gahan'        ; UserPrincipalName = 'dave.gahan@example.com'         ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Vocalists')                  ; Title = 'Lead Vocalist'               ; Email = 'dave.gahan@example.com'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0021' ; MobilePhone = '+44 7700 444442' ; Street = '32 Abbey Lane'            ; City = 'London'      ; PostalCode = 'EC2 1AA'     ; Company = 'Example Music Ltd'   ; Manager = 'Martin Gore'      ; Description = 'Lead vocalist for Depeche Mode'                               },
      @{ Name = 'Martin Gore'              ; SamAccountName = 'martin.gore'       ; UserPrincipalName = 'martin.gore@example.com'        ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Guitarists','Keyboards')     ; Title = 'Guitarist/Keyboardist'       ; Email = 'martin.gore@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0022' ; MobilePhone = '+44 7700 444441' ; Street = '32 Abbey Lane'            ; City = 'London'      ; PostalCode = 'EC2 1AA'     ; Company = 'Example Music Ltd'   ; Manager = ''                 ; Description = 'Guitarist, keyboardist & primary songwriter for Depeche Mode' },
      @{ Name = 'Alan Wilder'              ; SamAccountName = 'alan.wilder'       ; UserPrincipalName = 'alan.wilder@example.com'        ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Keyboards','Percussion')     ; Title = 'Keyboardist/Drummer'         ; Email = 'alan.wilder@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0023' ; MobilePhone = '+44 7700 444444' ; Street = '32 Abbey Lane'            ; City = 'London'      ; PostalCode = 'EC2 1AA'     ; Company = 'Example Music Ltd'   ; Manager = 'Martin Gore'      ; Description = 'Multi-instrumentalist for Depeche Mode (1982-1995, departed)' },
      @{ Name = 'Andrew Fletcher'          ; SamAccountName = 'andrew.fletcher'   ; UserPrincipalName = 'andrew.fletcher@example.com'    ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Keyboards')                  ; Title = 'Keyboards/Bass Synth'        ; Email = 'andrew.fletcher@example.com'      ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0024' ; MobilePhone = '+44 7700 444443' ; Street = '32 Abbey Lane'            ; City = 'London'      ; PostalCode = 'EC2 1AA'     ; Company = 'Example Music Ltd'   ; Manager = 'Martin Gore'      ; Description = 'Keyboard and bass synthesizer for Depeche Mode (deceased)'    },

      ## ========== TV-2 (Danmark/Copenhagen) ==========
      @{ Name = 'Steffen Brandt'           ; SamAccountName = 'steffen.brandt'    ; UserPrincipalName = 'steffen.brandt@example.com'     ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Vocalists','Guitarists')             ; Title = 'Lead Vocalist / Guitarist'   ; Email = 'steffen.brandt@example.com'       ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 0000 2222'    ; MobilePhone = '+45 50 12 3457'  ; Street = '1 Raadhuspladsen'         ; City = 'Copenhagen'  ; PostalCode = '1550'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Frontman of TV-2'                                             },
      @{ Name = 'Hans Erik Lerchenfeldt'   ; SamAccountName = 'hans.lerchenfeldt' ; UserPrincipalName = 'hans.lerchenfeldt@example.com'  ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Musicians')                          ; Title = 'Bassist'                     ; Email = 'hans.lerchenfeldt@example.com'    ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 3312 3457'    ; MobilePhone = '+45 20 11 1157'  ; Street = 'Nørrebrogade 2'           ; City = 'Copenhagen'  ; PostalCode = '2200'        ; Company = 'Example Music ApS'   ; Manager = 'Steffen Brandt'   ; Description = 'Bassist for TV-2'                                             },
      @{ Name = 'Sven Gaul'                ; SamAccountName = 'sven.gaul'         ; UserPrincipalName = 'sven.gaul@example.com'          ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Musicians')                          ; Title = 'Drummer'                     ; Email = 'sven.gaul@example.com'            ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 3312 3458'    ; MobilePhone = '+45 20 11 1158'  ; Street = 'Nørrebrogade 3'           ; City = 'Copenhagen'  ; PostalCode = '2200'        ; Company = 'Example Music ApS'   ; Manager = 'Steffen Brandt'   ; Description = 'Drummer for TV-2'                                             },
      @{ Name = 'Georg Olesen'             ; SamAccountName = 'georg.olesen'      ; UserPrincipalName = 'georg.olesen@example.com'       ; OU = @('Locations','Danmark','Nord Jyland', 'Aarhus','TV-2')                          ; Groups = @('TV-2','Musicians','Former Staff')           ; Title = 'Guitarist (Former)'          ; Email = 'georg.olesen@example.com'         ; Country = 'DK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aarhus Office'     ; Phone = '+45 8612 3470'    ; MobilePhone = '+45 20 11 1170'  ; Street = 'Åboulevarden 20'          ; City = 'Aarhus'      ; PostalCode = '8000'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Former guitarist and co-founder of TV-2 (1981-2003)'          },

      ## ========== Rocazino (Danmark/Koge) ========  ==
      @{ Name = 'Ulla Kjaer'               ; SamAccountName = 'ulla.kjaer'        ; UserPrincipalName = 'ulla.kjaer@example.com'         ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Vocalists')                      ; Title = 'Lead Vocalist'               ; Email = 'ulla.kjaer@example.com'           ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2234'    ; MobilePhone = '+45 3012 3456'   ; Street = '7 Torvet'                 ; City = 'Koge'        ; PostalCode = '4600'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Lead vocalist for Rocazino'                                   },
      @{ Name = 'Michael Bruun'            ; SamAccountName = 'michael.bruun'     ; UserPrincipalName = 'michael.bruun@example.com'      ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Guitarists')                     ; Title = 'Guitarist'                   ; Email = 'michael.bruun@example.com'        ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2235'    ; MobilePhone = '+45 3012 3457'   ; Street = '7 Torvet'                 ; City = 'Koge'        ; PostalCode = '4600'        ; Company = 'Example Music ApS'   ; Manager = 'Ulla Kjaer'       ; Description = 'Guitarist and songwriter for Rocazino'                        },
      @{ Name = 'Jan Sivertsen'            ; SamAccountName = 'jan.sivertsen'     ; UserPrincipalName = 'jan.sivertsen@example.com'      ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Percussion')                     ; Title = 'Drummer'                     ; Email = 'jan.sivertsen@example.com'        ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2236'    ; MobilePhone = '+45 3012 3458'   ; Street = '7 Torvet'                 ; City = 'Koge'        ; PostalCode = '4600'        ; Company = 'Example Music ApS'   ; Manager = 'Ulla Kjaer'       ; Description = 'Drummer for Rocazino'                                         },

      ## ========== Whigfield (Danmark/Kørsor og Faxe) ==========
      @{ Name = 'Sannie Charlotte Carlson' ; SamAccountName = 'whigfield'         ; UserPrincipalName = 'whigfield@example.com'          ; OU = @('Locations','Danmark','Sjælland','Køge','Kørsor','Whigfield','Users')          ; Groups = @('Whigfield','Warehouse Access')              ; Title  = 'Performer'                  ; Email  = 'whigfield@example.com'           ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Kørsor Office'     ; Phone = '+45 0000 3235'    ; MobilePhone = '+45 4012 3457'   ; Street = 'Torvegade 35.'            ; City = 'Faxe'        ; PostalCode = '4640'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Danish singer-songwriter'                                     },

      ## ========== MØ (Danmark/Odense) ==========
      @{ Name = 'Karen Marie Orsted'       ; SamAccountName = 'karen.orsted'      ; UserPrincipalName = 'mo@example.com'                 ; OU = @('Locations','Danmark','Fyn', 'Odense','Mo')                                    ; Groups = @('Mo','Vocalists')                            ; Title = 'Singer / Songwriter'         ; Email = 'mo@example.com'                   ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Odense Office'     ; Phone = '+45 0000 3234'    ; MobilePhone = '+45 4012 3456'   ; Street = '22 Vestergade'            ; City = 'Odense'      ; PostalCode = '5000'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Danish singer-songwriter known internationally as MØ'         },

      ## ========== Lis Sørensen (Danmark/Bramming) ==========
      @{ Name = 'Lis Sørensen'             ; SamAccountName = 'lissorensen'       ; UserPrincipalName = 'lis.sorensen@example.com'       ; OU = @('Locations','Danmark','Syd Jyland','Bramming','Music','Lis Sørensen','Users')  ; Groups  = @('Lis Sørensen','Studio Access')             ; Title   = 'Singer'                    ; Email   = 'lis.sorensen@example.com'       ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bramming Studio'   ; Phone = '+45 0000 4411'    ; MobilePhon  = '+45 4099 8822'   ; Street = 'Nørregade 12'             ; City = 'Bramming'    ; PostalCode = '1165'        ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Danish singer; known for the song Brændt'                     },

      ## ========== Helena Christensen (Danmark/Esjberg) ==========
      @{ Name = 'Helena Christensen'       ; SamAccountName = 'helenachristensen' ; UserPrincipalName = 'helena.christensen@example.com' ; OU = @('Locations','Danmark','Syd Jyland','Esbjerg','Creative','Photography','Users') ; Groups = @('Creative Team', 'Esbjerg Office Access')    ;  Title = 'Company Photographer'       ;  Email  = 'helena.christensen@example.com' ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ;  Department = 'Art'  ; Office = 'Esjberg Studio'    ; Phone = '+45 00 007 788'   ; MobilePhone = '+45 4011 2233'   ; Street = 'Skolegade 18'             ; City = 'Esjberg'     ; PostalCode = '6700'        ; Company = 'Example Creative ApS'; Manager = ''                 ; Description = 'Model/Photographer working on Chris Issak Wicked Game video'  },

      ## ========== Kraftwerk (West Germany/Bonn) ==========
      @{ Name = 'Ralf Hutter'              ; SamAccountName = 'ralf.hutter'       ; UserPrincipalName = 'ralf.hutter@example.net'        ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Vocalists','Musicians')         ; Title = 'Vocals/Synthesizer'          ; Email = 'ralf.hutter@example.net'          ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 111'  ; MobilePhone = '+49 1701 1111'   ; Street = 'Adenauerallee 1'          ; City = 'Bonn'        ; PostalCode = '53113'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Co-founder and frontman of Kraftwerk'                         },
      @{ Name = 'Florian Schneider'        ; SamAccountName = 'florian.schneider' ; UserPrincipalName = 'florian.schneider@example.net'  ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians')                     ; Title = 'Synthesizer/Flute'           ; Email = 'florian.schneider@example.net'    ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 112'  ; MobilePhone = '+49 1701 1112'   ; Street = 'Adenauerallee 2'          ; City = 'Bonn'        ; PostalCode = '53113'       ; Company = 'Example Music GmbH'  ; Manager = 'Ralf Hutter'      ; Description = 'Co-founder of Kraftwerk (1947-2020) - Account disabled'       },
      @{ Name = 'Wolfgang Flur'            ; SamAccountName = 'wolfgang.flur'     ; UserPrincipalName = 'wolfgang.flur@example.net'      ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians','Percussionists')    ; Title = 'Electronic Drums'            ; Email = 'wolfgang.flur@example.net'        ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 113'  ; MobilePhone = '+49 1701 1113'   ; Street = 'Adenauerallee 3'          ; City = 'Bonn'        ; PostalCode = '53113'       ; Company = 'Example Music GmbH'  ; Manager = 'Ralf Hutter'      ; Description = 'Electronic percussionist for Kraftwerk'                       },
      @{ Name = 'Karl Bartos'              ; SamAccountName = 'karl.bartos'       ; UserPrincipalName = 'karl.bartos@example.net'        ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians','Percussionists')    ; Title = 'Electronic Percussion'       ; Email = 'karl.bartos@example.net'          ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 114'  ; MobilePhone = '+49 1701 1114'   ; Street = 'Adenauerallee 4'          ; City = 'Bonn'        ; PostalCode = '53113'       ; Company = 'Example Music GmbH'  ; Manager = 'Ralf Hutter'      ; Description = 'Electronic percussionist and composer for Kraftwerk'          },

      ## ========== Fancy - Germany/Munich ==========
      @{ Name = 'Manfred Segieth'         ; SamAccountName = 'fancy'              ; UserPrincipalName = 'fancy@example.net'              ; OU = @('Locations','Germany','Bayern','Munich','Fancy')                               ; Groups = @('Fancy Solo','Vocalists')                    ; Title = 'Solo Artist'                 ; Email = 'fancy@example.net'                ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Munich Office'     ; Phone = '+49 89 3333 111'  ; MobilePhone = '+49 1723 3111'   ; Street = 'Leopoldstrasse 100'       ; City = 'Munich'      ; PostalCode = '80802'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Solo artist aus Munich, known for Lady of Ice & Italo disco'  },

      ## ==========  Nena - Germany/West Berlin ==========
      @{ Name = 'Gabriele Kerner'          ; SamAccountName = 'nena'              ; UserPrincipalName = 'nena@example.net'               ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Vocalists')                          ; Title = 'Lead Vocalist'               ; Email = 'nena@example.net'                 ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 111'  ; MobilePhone = '+49 1712 2111'   ; Street = 'Kurfurstendamm 100'       ; City = 'Berlin'      ; PostalCode = '10709'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Lead vocalist of Nena, known as Nena'                         },
      @{ Name = 'Carlo Karges'             ; SamAccountName = 'carlo.karges'      ; UserPrincipalName = 'carlo.karges@example.net'       ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Guitarist'                   ; Email = 'carlo.karges@example.net'         ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 112'  ; MobilePhone = '+49 1712 2112'   ; Street = 'Kurfurstendamm 101'       ; City = 'Berlin'      ; PostalCode = '10709'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Guitarist for Nena (1951-2002) - Account disabled'            },
      @{ Name = 'Uwe Fahrenkrog-Petersen'  ; SamAccountName = 'uwe.fahrenkrog'    ; UserPrincipalName = 'uwe.fahrenkrog@example.net'     ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Keyboardist'                 ; Email = 'uwe.fahrenkrog@example.net'       ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 113'  ; MobilePhone = '+49 1712 2113'   ; Street = 'Kurfurstendamm 102'       ; City = 'Berlin'      ; PostalCode = '10709'       ; Company = 'Example Music GmbH'  ; Manager = 'Gabriele Kerner'  ; Description = 'Keyboardist and songwriter for Nena'                          },
      @{ Name = 'Jurgen Dehmel'            ; SamAccountName = 'jurgen.dehmel'     ; UserPrincipalName = 'jurgen.dehmel@example.net'      ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Bassist'                     ; Email = 'jurgen.dehmel@example.net'        ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 114'  ; MobilePhone = '+49 1712 2114'   ; Street = 'Kurfurstendamm 103'       ; City = 'Berlin'      ; PostalCode = '10709'       ; Company = 'Example Music GmbH'  ; Manager = 'Gabriele Kerner'  ; Description = 'Bassist for Nena'                                             },
      @{ Name = 'Rolf Brendel'             ; SamAccountName = 'rolf.brendel'      ; UserPrincipalName = 'rolf.brendel@example.net'       ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians','Percussionists')         ; Title = 'Drummer'                     ; Email = 'rolf.brendel@example.net'         ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 115'  ; MobilePhone = '+49 1712 2115'   ; Street = 'Kurfurstendamm 104'       ; City = 'Berlin'      ; PostalCode = '10709'       ; Company = 'Example Music GmbH'  ; Manager = 'Gabriele Kerner'  ; Description = 'Drummer for Nena'                                             },

      ## ========== Tangerine Dream - Germany/West Berlin ==========
      @{ Name = 'Edgar Froese'            ; SamAccountName = 'edgar.froese'       ; UserPrincipalName = 'edgar.froese@example.net'       ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer/Guitar'          ; Email = 'edgar.froese@example.net'         ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 121'  ; MobilePhone = '+49 171 2221121' ; Street = 'Kantstrasse 50'           ; City = 'Berlin'      ; PostalCode = '10625'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Founder of Tangerine Dream (1944-2015) - Account disabled'    },
      @{ Name = 'Christopher Franke'      ; SamAccountName = 'christopher.franke' ; UserPrincipalName = 'christopher.franke@example.net' ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer/Drums'           ; Email = 'christopher.franke@example.net'   ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 122'  ; MobilePhone = '+49 171 2221122' ; Street = 'Kantstrasse 51'           ; City = 'Berlin'      ; PostalCode = '10625'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Synthesizer and electronic drums for Tangerine Dream'         },
      @{ Name = 'Peter Baumann'           ; SamAccountName = 'peter.baumann'      ; UserPrincipalName = 'peter.baumann@example.net'      ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer'                 ; Email = 'peter.baumann@example.net'        ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 123'  ; MobilePhone = '+49 171 2221123' ; Street = 'Kantstrasse 52'           ; City = 'Berlin'      ; PostalCode = '10625'       ; Company = 'Example Music GmbH'  ; Manager = ''                 ; Description = 'Synthesizer player for Tangerine Dream'                       },

      ## Dummy user- for setting up new users via template - as is common
      @{ Name = 'Template User'           ; SamAccountName = 'template.users'     ; UserPrincipalName = 'template.user@example.com'      ; OU = @('Locations','Templates','Users')                                               ; Groups = @('Disabled Users','Template Users')           ; Title = 'Template User'               ; Email = 'template.user@example.com'        ; Country = 'GL' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+00 00 0000 000'  ; MobilePhone = '+00 000 0000000' ; Street = '1 Example Street'         ; City = 'Nuuk'        ; PostalCode = '00000'       ; Company = 'Example Music ApS'   ; Manager = ''                 ; Description = 'Counting Paperclips'                                          }
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
      @{ Name = 'The Proclaimers'       ; Description = 'Scottish folk rock duo formed in Edinburgh in 1983'            ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Craig Reid'         ; Email='proclaimers@example.net'      },

      ## Occupation groups. Normally used for distirb lists in Exchange (365)
      @{ Name = 'Vocalists'             ; Description = 'Lead singers and vocalists across all bands'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'vocalists@example.com'      },
      @{ Name = 'Keyboards'             ; Description = 'Keyboard And synthesizers'                                     ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'synths@example.com'         },
      @{ Name = 'Musicians'             ; Description = 'Instrumentalists'                                              ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'instruments@example.com'    },
      @{ Name = 'Guitarists'            ; Description = 'Guitar and bass players'                                       ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'guitars@example.com'        },
      @{ Name = 'Percussionists'        ; Description = 'Drummers and percussion specialists'                           ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'percussion@example.com'     },

      ## Non band based groups
      @{ Name = 'Sales'                 ; Description = 'Remote VPN users and road warriors in sales'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy ='IT Admin'            ; Email='sales@example.com'            },
      @{ Name = 'Former Staff'          ; Description = 'Former band members and staff'                                 ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = ''                           },
      @{ Name = 'VPN-Users'             ; Description='Remote access users (touring staff, laptops and tablets)'        ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'IT Operations'      ; Email='vpn-users@example.net'        },
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
      @{ Name = 'EXAPHNABD001' ; SamAccountName = 'EXAPHNABD001$' ; Type = 'Phones'        ; Role = 'PHN' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAPHNABD001.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Devices')           ; OS = 'iOS'                               ; OperatingSystemVersion = '17.x'                 ; Description = 'Corporate iPhone – Annie Lennox'                ; LastLogon = 'N/A'                      ; IPAddress = 'N/A'             ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  };
      @{ Name = 'EXAPHNABD002' ; SamAccountName = 'EXAPHNABD002$' ; Type = 'Phones'        ; Role = 'PHN' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAPHNABD002.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Devices')           ; OS = 'iOS'                               ; OperatingSystemVersion = '17.x'                 ; Description = 'Corporate iPhone – Dave Stewart'                ; LastLogon = 'N/A'                      ; IPAddress = 'N/A'             ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXAMBPABD001' ; SamAccountName = 'EXAMBPABD001$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAMBPABD001.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Computers')         ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '14.2.1'               ; Description = 'MacBook – Annie Lennox (Aberdeen satellite)'    ; LastLogon = (Get-Date).AddDays(-16)    ; IPAddress = '192.168.224.137' ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXAMBPABD002' ; SamAccountName = 'EXAMBPABD002$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAMBPABD002.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Computers')         ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '14.2.1'               ; Description = 'MacBook – Dave Stewart (Aberdeen satellite)'    ; LastLogon = (Get-Date).AddDays(-17)    ; IPAddress = '192.168.224.124' ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

      ## Edinburgh, Scotland
      @{ Name = 'EXAWKSEDI001' ; SamAccountName = 'EXAWKSEDI001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXAWKSEDI001.example.org' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')        ; OS = 'Windows 10 Pro'                    ; OperatingSystemVersion = '22H2'                 ; Description = 'Shared desktop workstation'                     ; LastLogon = (Get-Date).AddDays(-17)    ; IPAddress = '192.168.131.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Edi#Wks!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(3).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-27) ; LAPSPasswordExpiration = (Get-Date).AddDays(33) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXALAPEDI098' ; SamAccountName = 'EXALAPEDI098$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXALAPEDI098.example.com' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')        ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Pool laptop – Edinburgh Office'                 ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.131.108' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Edi#Soon!98'    ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(1).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-25) ; LAPSPasswordExpiration = (Get-Date).AddDays(55) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXARACEDI001' ; SamAccountName = 'EXARACEDI001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXARACEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Infrastructure')   ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / calvin'                          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASWIEDI001' ; SamAccountName = 'EXASWIEDI001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASWIEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Network')          ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Floor switch – admin:cisco / EDI2960!'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASWIEDI002' ; SamAccountName = 'EXASWIEDI001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASWIEDI002.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Network')          ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port - admin:cisco123'        ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.6'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASBCEDI001' ; SamAccountName = 'EXASBCEDI001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASBCEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Servers')          ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-EDI-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      ## Internet Connected cofee pot https://en.wikipedia.org/wiki/Trojan_Room_coffee_pot?useskin=vector
      @{ Name = 'EXATEAEDI001' ; SamAccountName = 'EXATEAEDI001$' ; Type = 'Siemens EQ700' ; Role = 'TEA' ; Site = 'EDI' ; Location = 'Edinburgh Office Break Room'  ; DNSHostName = 'EXATEAEDI001.example.com' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')        ; OS = 'Smart Bean to Cup Coffee Machine'  ; OperatingSystemVersion = 'TP713GB9'             ; Description = 'Coffee Machine in Edinburgh office kitchen'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.16'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

      ## From Miami to Canada B-)
      @{ Name = 'EXALAPBRK001' ; SamAccountName = 'EXALAPBRK001$' ; Type='Computer'        ; Role='LAP'   ; Site = 'BRK'  ; Location = 'Brockille, ON, CA'           ; DNSHostName = 'EXALAPBRK001.example.net' ; OU = @('Locations','CA','Ontario','Brockville','The Proclaimers')  ; OS='Windows 11 Pro'                      ; OperatingSystemVersion = '24H2'                 ; Description = 'Tour laptop (Canada)'                           ; LastLogon = (Get-Date).AddDays(-11)    ; IPAddress = '10.20.10.21'     ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@18'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-15) ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXALAPMIA001' ; SamAccountName = 'EXALAPMIA001$' ; Type='Computer'        ; Role='LAP'   ; Site = 'MIA'  ; Location = 'Miami, FL, US'               ; DNSHostName = 'EXALAPMIA001.example.net' ; OU = @('Locations','US','Florida','Miami','The Proclaimers'     )  ; OS='macOS Sonoma'                        ; OperatingSystemVersion = '25.2'                 ; Description = 'Tour MacBook (Miami)'                           ; LastLogon = (Get-Date).AddDays(-14)    ; IPAddress = '10.30.10.21'     ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@18'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-15) ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },

      ## Glasgow
      @{ Name = 'EXAWKSGLA001' ; SamAccountName = 'EXAWKSGLA001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAWKSGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')          ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Glasgow office'          ; LastLogon = (Get-Date).AddHours(-2)    ; IPAddress = '192.168.141.150' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Wks!41'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(9).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-12) ; LAPSPasswordExpiration = (Get-Date).AddDays(9)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXAWKSGLA002' ; SamAccountName = 'EXAWKSGLA002$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAWKSGLA002.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')          ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Glasgow office'          ; LastLogon = (Get-Date).AddDays(-1)     ; IPAddress = '192.168.141.151' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Wks!73'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(14).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-8)  ; LAPSPasswordExpiration = (Get-Date).AddDays(14) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXALAPGLA001' ; SamAccountName = 'EXALAPGLA001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXALAPGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')          ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Dell Latitude laptop - Pool device'             ; LastLogon = (Get-Date).AddHours(-5)    ; IPAddress = '192.168.141.152' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Lap@19'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(6).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-10) ; LAPSPasswordExpiration = (Get-Date).AddDays(6)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXAPRNGLA001' ; SamAccountName = 'EXAPRNGLA001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAPRNGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')          ; OS = 'Printer'                           ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP LaserJet Pro - Main floor'                   ; LastLogon = (Get-Date).AddMinutes(-30) ; IPAddress = '192.168.141.16'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

      ## Avocado Central
      @{ Name = 'EXAWKSLND001' ; SamAccountName = 'EXAWKSLND001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'LND' ; Location = 'London, England'              ; DNSHostName = 'EXAWKSLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')            ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - London office'           ; LastLogon = (Get-Date).AddHours(-1)    ; IPAddress = '192.168.20.150'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Bon#Wks!22'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(10).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-14) ; LAPSPasswordExpiration=(Get-Date).AddDays(10)   ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXAPRNLND001' ; SamAccountName = 'EXAPRNLND001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'LND' ; Location = 'London, England'              ; DNSHostName = 'EXAPRNLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')            ; OS = 'Printer'                           ; OperatingSystemVersion = 'N/A'                  ; Description = 'Xerox WorkCentre - Reception'                   ; LastLogon = (Get-Date).AddMinutes(-15) ; IPAddress = '192.168.20.16'   ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXARTRLND001' ; SamAccountName = 'EXARTRLND001$' ; Type = 'Network'       ; Role = 'RTR' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXARTRLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')            ; OS = 'Cisco ISR 4331'                    ; OperatingSystemVersion = 'ISR4331-LND-552901'   ; Description = 'WAN edge router'                                ; LastLogon = (Get-Date).AddHours(-13)   ; IPAddress = '192.168.20.254'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXAFWLLND001' ; SamAccountName = 'EXAFWLLND001$' ; Type = 'Network'       ; Role = 'FWL' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXAFWLLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')            ; OS = 'Cisco ASA 5516-X'                  ; OperatingSystemVersion = 'ASA5516-LND-884210'   ; Description = 'Perimeter firewall and VPN gateway'             ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.20.1'    ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXARACLND001' ; SamAccountName = 'EXARACLND001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXARACLND001.example.net' ; OU = @('Locations','UK','England','London','Infrastructure')       ; OS ='Dell iDRAC9'                        ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell PowerEdge iDRAC – root / P@ssLND-RAC01'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASWILND001' ; SamAccountName = 'EXASWILND001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXASWILND001.example.net' ; OU = @('Locations','UK','England','London','Network')              ; OS ='Cisco Catalyst 9300'                ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core switch – admin:cisco / Sw1tchLND!'         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASBCLND001' ; SamAccountName = 'EXASBCLND001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXASBCLND001.example.net' ; OU = @('Locations','UK','England','London','Servers')              ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-LND-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

      ## Munich, DE
      @{ Name = 'EXAWKSMUN001' ; SamAccountName = 'EXAWKSMUN001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'BON' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAWKSMUN001.example.net' ; OU = @('Locations','Germany','Munich','Computers')                 ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Bonn office'             ; LastLogon = (Get-Date).AddHours(-3)    ; IPAddress = '192.168.89.150'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Wks!41'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(11).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-19) ; LAPSPasswordExpiration = (Get-Date).AddDays(11) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXALAPMUN001' ; SamAccountName = 'EXALAPMUN001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXALAPMUN001.example.net' ; OU = @('Locations','Germany','Munich','Computers')                 ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Pool laptop – Munich office'                    ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.89.151'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Lap!47'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(2).ToFileTimeUtc()  ; LAPSPasswordLastSet=(Get-Date).AddDays(-28)   ; LAPSPasswordExpiration = (Get-Date).AddDays(2)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXALAPMUN002' ; SamAccountName = 'EXALAPMUN002$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXALAPMUN002.example.net' ; OU = @('Locations','Germany','Munich','Computers')                 ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Pool laptop – Munich office'                    ; LastLogon = (Get-Date).AddDays(-95)    ; IPAddress = '192.168.89.152'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Lap!02'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(-61).ToFileTimeUtc(); LAPSPasswordLastSet=(Get-Date).AddDays(-91)   ; LAPSPasswordExpiration = (Get-Date).AddDays(-61); TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
      @{ Name = 'EXARACMUN001' ; SamAccountName = 'EXARACMUN001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAracMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Infrastructure')            ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – admin / MUN-ilo-pass'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASWIMUN001' ; SamAccountName = 'EXASWIMUN001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAswiMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Network')                   ; OS = 'Cisco Catalyst 9200'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Access switch – admin:cisco / MUN9200!'         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
      @{ Name = 'EXASBCMUN001' ; SamAccountName = 'EXASBCMUN001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAsbcMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Servers')                   ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-MUN-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

      ## Bonn, DE
      @{ Name = 'EXALAPBON001' ; SamAccountName = 'EXALAPBON001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'BON' ; Location = 'Bonn, Germany'                ; DNSHostName = 'EXALAPBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Lenovo ThinkPad - Pool device'                  ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.228.150' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Bon#Lap@64'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(5).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-18) ; LAPSPasswordExpiration = (Get-Date).AddDays(5)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $false }, ## Disabled for maintenance
      @{ Name = 'EXAWKSBON001' ; SamAccountName = 'EXAWKSBON001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAWKSBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '23H2'                 ; Description = 'Finance department desktop workstation'         ; LastLogon = (Get-Date).AddDays(-11)    ; IPAddress = '192.168.228.151' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'B0n#Wks!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(12).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-20) ; LAPSPasswordExpiration = (Get-Date).AddDays(10) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAWKSBON002' ; SamAccountName = 'EXAWKSBON002$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAWKSBON002.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Finance Department Workstation'                 ; LastLogon = (Get-Date).AddHours(-2)    ; IPAddress = '192.168.228.152' ; Domain = 'example.net' ;  LAPSPassword = 'Kx9#mP2$vL5@qR8!'     ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; LAPSPasswordLastSet  = (Get-Date).AddDays(-15) ; 'ms-Mcs-AdmPwd' = 'Kx9#mP2$vL5@qR8!' ; 'ms-Mcs-AdmPwdExpirationTime' = '133789234567890123'                    ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  ; BitLockerRecoveryKey = '648392-174853-293847-483920-756483-029384-475829-384756' }, ## There's always that one guy
      @{ Name = 'EXALAPBON002' ; SamAccountName = 'EXALAPBON002$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXALAPBON002.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '23H2'                 ; Description = 'Assigned laptop for finance staff'              ; LastLogon = (Get-Date).AddDays(-21)    ; IPAddress = '192.168.228.153' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'B0n#Lap!02'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-14) ; LAPSPasswordExpiration = (Get-Date).AddDays(7)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAVCUBON001' ; SamAccountName = 'EXAVCUBON001$' ; Type = 'AV'            ; Role = 'VCU' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAVCUBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Poly Studio X70'                   ; OperatingSystemVersion = 'POLY-X70-BON-772190'  ; Description = 'Boardroom video conferencing system'            ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.228.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXACAMBON003' ; SamAccountName = 'EXACAMBON003$' ; Type = 'Security'      ; Role = 'CAM' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXACAMBON003.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Axis P3245-LVE'                    ; OperatingSystemVersion = 'AXIS3245-BON-009381'  ; Description = 'CCTV security camera'                           ; LastLogon = (Get-Date).AddDays(-12)    ; IPAddress = '192.168.228.17'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXATVSBON001' ; SamAccountName = 'EXATVSBON001$' ; Type = 'AV'            ; Role = 'TVS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXATVSBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                   ; OS = 'Samsung Business TV 65"'           ; OperatingSystemVersion = 'SAMS65-BON-771992'    ; Description = 'Networked information display'                  ; LastLogon = (Get-Date).AddDays(-13)    ; IPAddress = '192.168.151.18'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACBON001' ; SamAccountName = 'EXARACBON001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXARACBON001.example.net' ; OU = @('Locations','Germany','Bonn','Infrastructure')              ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / BON-RAC-01'                      ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWIBON001' ; SamAccountName = 'EXASWIBON001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXASWIBON001.example.net' ; OU = @('Locations','Germany','Bonn','Network')                     ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Office switch – admin:cisco / BONsw01'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCBON001' ; SamAccountName = 'EXASBCBON001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXASBCBON001.example.net' ; OU = @('Locations','Germany','Bonn','Servers')                     ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-BON-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

     ## Liverpool, UK
      @{ Name = 'EXASVRLIV001' ; SamAccountName = 'EXASVRLIV001$' ; Type = 'Server'        ; Role = 'SVR' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASVRLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')         ; OS = 'Windows Server 2022'               ; OperatingSystemVersion = '21H2'                 ; Description = 'File Server for Liverpool site'                 ; LastLogon = (Get-Date).AddDays(-8)     ; IPAddress = '192.168.151.10'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = '$3s$4m3BuN*'    ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-1)  ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAMBPLIV001' ; SamAccountName = 'EXAMBPLIV001$' ; Type = 'Macbook'       ; Role = 'MBP' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXAMBPLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')         ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Macbook Pro 2024 - Pool device'                 ; LastLogon = (Get-Date).AddDays(-14)    ; IPAddress = '192.168.151.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Liv#Mac@31'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(13).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(13) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }, ## If TPM is disabled we don't have the other stuff
      @{ Name = 'EXAMACLIV001' ; SamAccountName = 'EXAMACLIV001$' ; Type = 'iMac'          ; Role = 'MAC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXAMACLIV001.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')         ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Macbook Pro 2024 - Pool device'                 ; LastLogon = (Get-Date).AddDays(-40)    ; IPAddress = '192.168.151.152' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Liv#Mac@77'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(17).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-35) ; LAPSPasswordExpiration = (Get-Date).AddDays(17) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $false }, ## Disabled for maintenance
      @{ Name = 'EXARDRLIV002' ; SamAccountName = 'EXARDRLIV002$' ; Type = 'Security'      ; Role = 'RDR' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXARDRLIV002.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')         ; OS = 'HID Signo Reader'                  ; OperatingSystemVersion = 'HID-SIGNO-LIV-338572' ; Description = 'Electronic badge reader for access control'     ; LastLogon = (Get-Date).AddDays(-9)     ; IPAddress = '192.168.31.16'   ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXABPSLIV001' ; SamAccountName = 'EXABPSLIV001$' ; Type = 'Security'      ; Role = 'BPS' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXABPSLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')         ; OS = 'HID Omnikey 5427'                  ; OperatingSystemVersion = 'OMNI5427-LIV-449821'  ; Description = 'Badge programming workstation'                  ; LastLogon = (Get-Date).AddDays(-23)    ; IPAddress = '192.168.151.17'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACLIV001' ; SamAccountName = 'EXARACLIV001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXARACLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Infrastructure')    ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – admin / liv-ilo-pass'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWILIV001' ; SamAccountName = 'EXASWILIV001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASWILIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Network')           ; OS = 'Cisco Catalyst 9200'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core switch – admin:cisco / LIVcore!'           ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCLIV001' ; SamAccountName = 'EXASBCLIV001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASBCLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-LIV-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Newcastle, UK
      @{ Name = 'EXASRVNEW001' ; SamAccountName = 'EXASRVNEW001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASRVNEW001.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')         ; OS = 'Windows Server 2022'               ; OperatingSystemVersion = '21H2'                 ; Description = 'File and print server for Newcastle office'     ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.191.21'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'New#Srv!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(13) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXARACNEW001' ; SamAccountName = 'EXARACNEW001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXARACNEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Infrastructure')    ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / new-rac-01!'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWINEW001' ; SamAccountName = 'EXASWINEW001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASWINEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Network')           ; OS = 'TP-Link JetStream'                 ; OperatingSystemVersion = 'N/A'                  ; Description = 'Access switch – admin / NEWsw!'                 ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCNEW001' ; SamAccountName = 'EXASBCNEW001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASBCNEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-NEW-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAWKSNEW099' ; SamAccountName = 'EXAWKSNEW099$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXAWKSNEW099.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')         ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Newcastle office PC'                            ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.191.161' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'New#Expired!99' ; 'msLAPS-PasswordExpirationTime'=(Get-Date).AddDays(-1).ToFileTimeUtc()   ; LAPSPasswordLastSet = (Get-Date).AddDays(-31) ; LAPSPasswordExpiration = (Get-Date).AddDays(-1) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      ## West Berlin, DE
      @{ Name = 'EXASRVBRD001' ; SamAccountName = 'EXASRVBRD001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'BRD' ; Location = 'West Berlin, FRG (DE)'        ; DNSHostName = 'EXASRVBRD001.example.net' ; OU = @('Locations','Germany','West Berlin','Computers')            ; OS = 'Windows Server 2019'               ; OperatingSystemVersion = '1809'                 ; Description = 'Legacy application server (West Berlin site)'   ; LastLogon = (Get-Date).AddDays(-1)     ; IPAddress = '192.168.30.21'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ged#Srv!19'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-25) ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXANIXBRD002' ; SamAccountName = 'EXANIXBRD002$' ; Type = 'Unix'          ; Role = 'NIX' ; Site = 'BRD' ; Location = 'West Berlin, FRG (DE)'        ; DNSHostName = 'EXANIXBRD002.example.net' ; OU = @('Locations','Germany','West Berlin','Computers')            ; OS = 'Debian 12'                         ; OperatingSystemVersion = '12.2'                 ; Description = 'Linux server hosting internal services'         ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.30.22'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Odense, DK
      @{ Name = 'EXAMACODE001' ; SamAccountName = 'EXAMACODE001$' ; Type = 'Computer'      ; Role = 'MAC' ; Site = 'ODE' ; Location = 'Odense, Danmark'              ; DNSHostName = 'EXAMACODE001.example.com' ; OU = @('Locations','Danmark','Odense','Computers')                 ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Design team iMac workstation'                   ; LastLogon = (Get-Date).AddDays(-7)     ; IPAddress = '192.168.66.150'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ode#Mac@44'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(14).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(14) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAMBPODE002' ; SamAccountName = 'EXAMBPODE002$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ODE' ; Location = 'Odense, Danmark'              ; DNSHostName = 'EXAMBPODE002.example.com' ; OU = @('Locations','Danmark','Odense','Computers')                 ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Executive MacBook Pro'                          ; LastLogon = (Get-Date).AddDays(-23)    ; IPAddress = '192.168.66.151'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ode#Mac@81'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(10).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-30) ; LAPSPasswordExpiration = (Get-Date).AddDays(10) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      ## Jukebox. This will be funny to you, if you have ever stayed at First Hotel Grand Odense Jernbanegade 18, 5000 Odense, Danmark
      @{ Name = 'EXAMUSODE001' ; SamAccountName = 'EXAMUSODE001$' ; Type = 'Jukebox'       ; Role = 'MUS' ; Site ='ODE'  ; Location = 'Odense Office – Ground Floor' ; DNSHostName = 'EXAMUSODE001.example.org' ; OU = @('Locations','Danmark','Odense','Computers')                 ; OS = 'Pureline 128V Retro Vinyl Jukebox' ; OperatingSystemVersion = 'PL128V-ODE-095823'    ; Description = 'Retro Vinyl jukebox in the Odense office.'      ; LastLogon = (Get-Date).AddHours(-4)    ; IPAddress = '192.168.66.18'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Køge, DK
      @{ Name = 'EXAWAPKGE001' ; SamAccountName = 'EXAWAPKGE001$' ; Type = 'Network'       ; Role = 'WAP' ; Site = 'KGE' ; Location = 'Køge, Danmark'                ; DNSHostName = 'EXAWAPKGE001.example.com' ; OU = @('Locations','Danmark','Køge','Computers')                   ; OS = 'Ubiquiti UniFi U6-Pro'             ; OperatingSystemVersion = 'U6P-KGE-847392'       ; Description = 'Enterprise Wi-Fi access point'                  ; LastLogon = (Get-Date).AddDays(-18)    ; IPAddress = '192.168.56.5'    ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAPRNKGE001' ; SamAccountName = 'EXAPRNKGE001$' ; Type = 'Peripheral'    ; Role = 'PRN' ; Site = 'KGE' ; Location = 'Køge, Danmark'                ; DNSHostName = 'EXAPRNKGE001.example.com' ; OU = @('Locations','Danmark','Køge','Computers')                   ; OS = 'HP LaserJet Enterprise MFP M528'   ; OperatingSystemVersion = 'HPM528-KGE-193847'    ; Description = 'Networked multifunction printer'                ; LastLogon = (Get-Date).AddDays(-10)    ; IPAddress = '192.168.56.16'   ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## København, DK
      @{ Name = 'EXACLKCPH001' ; SamAccountName = 'EXACLKCPH001$' ; Type = 'IoT'           ; Role = 'CLK' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXACLKCPH001.example.com' ; OU = @('Locations','Danmark','Copenhagen','Computers')             ; OS = 'Meinberg LANTIME M300'             ; OperatingSystemVersion = 'M300-CPH-661204'      ; Description = 'Network-synchronised NTP clock'                 ; LastLogon = (Get-Date).AddDays(-4)     ; IPAddress = '192.168.228.18'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACCPH001' ; SamAccountName = 'EXARACCPH001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXARACCPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Infrastructure')        ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / CPH-rac01!'                      ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWICPH001' ; SamAccountName = 'EXASWICPH001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXASWICPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Network')               ; OS = 'TP-Link JetStream'                 ; OperatingSystemVersion = 'N/A'                  ; Description = 'Office switch – admin / CPHsw!'                 ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCCPH001' ; SamAccountName = 'EXASBCCPH001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXASBCCPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Servers')               ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-CPH-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      @{ Name = 'EXATVSCPH001' ; SamAccountName = 'EXATVSCPH001$' ; Type = 'TVs'           ; Role = 'TVS' ; Site='CPH'   ; Location = 'København Centrum office'     ; DNSHostName = 'EXATVSCPH001.example.com' ; OU = @('Locations','Danmark','Copenhagen','Computers')             ; OS = 'Bella Kronik 42X'                  ; OperatingSystemVersion = 'BTV-042001'           ; Description = 'Bella TV streams DR/TV2 in the CPH office..'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.17'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Manchester, UK
      @{ Name = 'EXALAPMCR001' ; SamAccountName = 'EXALAPMCR001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXALAPMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')        ; OS = 'Windows 11 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Management laptop – Manchester office'          ; LastLogon = (Get-Date).AddDays(-12)    ; IPAddress = '192.168.228.19'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password'='Mcr#Lap@18'       ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-20) ; LAPSPasswordExpiration = (Get-Date).AddDays(7)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXALAPMCR002' ; SamAccountName = 'EXALAPMCR002$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXALAPMCR002.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')        ; OS = 'Windows 11 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Mobile staff laptop – Manchester office'        ; LastLogon = (Get-Date).AddDays(-24)    ; IPAddress = '192.168.161.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password'='Mcr#Lap@92'       ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(4).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-26) ; LAPSPasswordExpiration = (Get-Date).AddDays(4)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAWKSMCR001' ; SamAccountName = 'EXAWKSMCR001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAWKSMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')        ; OS = 'Windows 10 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Front desk workstation – Manchester office'     ; LastLogon = (Get-Date).AddDays(-40)    ; IPAddress = '192.168.161.152' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password'='Mcr#Wks!09'       ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(12).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-19) ; LAPSPasswordExpiration = (Get-Date).AddDays(12) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAWKSMCR002' ; SamAccountName = 'EXAWKSMCR002$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAWKSMCR002.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')        ; OS = 'Windows 10 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Finance workstation – Manchester office'        ; LastLogon = (Get-Date).AddDays(-9)     ; IPAddress = '192.168.161.153' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password'='Mcr#Wks!55'       ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(8).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-22) ; LAPSPasswordExpiration = (Get-Date).AddDays(8)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAPRNMCR001' ; SamAccountName = 'EXAPRNMCR001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAPRNMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')        ; OS = 'Embedded Printer Firmware'         ; OperatingSystemVersion = 'N/A'                  ; Description = 'Main network printer – Manchester office'       ; LastLogon = (Get-Date).AddDays(-4)     ; IPAddress = '192.168.161.16'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACMCR001' ; SamAccountName = 'EXARACMCR001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXARACMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Infrastructure')   ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – Administrator / mcr-ilo-01'               ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWIMCR001' ; SamAccountName = 'EXASWIMCR001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXASWIMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Network')          ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Distribution switch – admin:cisco / MCR9300!'   ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCMCR001' ; SamAccountName = 'EXASBCMCR001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXASBCMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Servers')          ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-MCR-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }

      ## Birmingham, UK
      @{ Name = 'EXASWIBIR001' ; SamAccountName = 'EXASWIBIR001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASWIBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')          ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core Switch'                                    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAFWLBIR001' ; SamAccountName = 'EXAFWLBIR001$' ; Type = 'Router'        ; Role = 'RTR' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAFWLBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')          ; OS = 'Palo Alto PanOS'                   ; OperatingSystemVersion = 'N/A'                  ; Description = 'Palo Alto Site Firewall / VPN Gateway'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.1'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASRVBIR001' ; SamAccountName = 'EXASRVBIR001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASRVBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Servers')          ; OS = 'Rocky Linux'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'Roxy Linux Node runnning Oracle DB'             ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAILOBIR001' ; SamAccountName = 'EXAILOBIR001$' ; Type = 'RAC'           ; Role = 'ILO' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXARACBIR001.example.net' ; OU = @('Locations','UK','England','Manchester','Infrastructure')   ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP iLO in Birmingham (Administrator:Ay3L0w@)'   ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCBIR001' ; SamAccountName = 'EXASBCBIR001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASBCBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Servers')          ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = '3CX Debian'           ; Description = '3CX SBC – ssh:root / 3cx-BIR-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAMBPBIR001' ; SamAccountName = 'EXAMBPBIR001$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAMBPBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Computers')        ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '25.2.1'               ; Description = 'MacBook Pro – touring'                          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.41'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXATABBIR001' ; SamAccountName = 'EXATABBIR001$' ; Type = 'Tablet'        ; Role = 'TAB' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXATABLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Tablets')           ; OS = 'Samsung Galaxy Tab A10'            ; OperatingSystemVersion = 'Lineage OS'           ; Description = 'Android tablet – setlists (service account)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.61'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAPHNBIR001' ; SamAccountName = 'EXAPHNBIR001$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAPHNLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Phones')            ; OS = 'Samsung Glalaxy S25 Ultra'         ; OperatingSystemVersion = 'Linage OS'            ; Description = 'Android phone – touring handset'                ; LastLogon = 'N/A'                      ; IPAddress = $null             ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWIBIR002' ; SamAccountName = 'EXASWIBIR002$' ; Type = 'Switch'        ; Role = 'SWI' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASWIBIR002.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')          ; OS = 'N/A'                               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port (admin:catalyst80s)'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACBIR001' ; SamAccountName = 'EXARACBIR001$' ; Type = 'RAC'           ; Role = 'RAC' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXARACBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Infrastructure')   ; OS = 'N/A'                               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell DRAC Birmingham   (dracadmin:Dr@c1983)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.6'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXALCDBIR001' ; SamAccountName = 'EXALCDBIR001$' ; Type = 'Wall Display'  ; Role = 'LCD' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXALCDBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Displays')         ; OS = 'NEC PlasmaSync 42MP1'              ; OperatingSystemVersion = 'PlasmaSync v1.x'      ; Description = 'LCD status display (NOC / call queue display)'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },


      ## Clydebank Infra
      @{ Name = 'EXASWICLY001' ; SamAccountName = 'EXASWICLY001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASWICLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'Cisco Catalyst 9300'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port (admin:catalyst80s)'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAFWLCLY001' ; SamAccountName = 'EXAFWLCLY001$' ; Type = 'Router'        ; Role = 'RTR' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXAFWLCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'FortiOS'                          ; OperatingSystemVersion = '7.6.5 build 3651'     ; Description = 'Firewall / VPN Gateway'                         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.1'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASRVCLY001' ; SamAccountName = 'EXASRVCLY001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASRVCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Servers')           ; OS = 'Rocky Linux'                      ; OperatingSystemVersion = 'N/A'                  ; Description = 'Database Server running Oracle DB'              ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXAILOCLY001' ; SamAccountName = 'EXAILOCLY001$' ; Type = 'RAC'           ; Role = 'ILO' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXARACCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Infrastructure')    ; OS = 'HPE iLO5'                         ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP iLO in Clydebank (Administrator:@nG3l3yE$)'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASBCCLY001' ; SamAccountName = 'EXASBCCLY001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASBCCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Servers')           ; OS = '3CX SBC Debian'                   ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-CLY-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASURCLY001' ; SamAccountName = 'EXASURCLY001$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASURCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')       ; OS = 'Windows 11 Pro'                   ; OperatingSystemVersion = '24H2'                 ; Description = 'Microsoft Surface – touring'                    ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.141.51'  ; Domain = 'example.net' ; 'msLAPS-AccountName' ='Administrator'  ; 'msLAPS-Password' ='Cly%SUR@07'      ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(22)                 ; LAPSPasswordLastSet = (Get-Date).AddDays(8)   ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
      @{ Name = 'EXAPHNGLA001' ; SamAccountName = 'EXAPHNCLY001$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXAPHNCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Phones') ;          ; OS = 'Apple iOS'                        ; OperatingSystemVersion = '26.1'                 ; Description = 'Android phone – touring handset'                ; LastLogon = 'N/A'                      ; IPAddress =  $null            ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXATABGLA001' ; SamAccountName = 'EXATABCLY001$' ; Type = 'Tablet'        ; Role = 'TAB' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXATABCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Computer')          ; OS = 'Apple iOD'                        ; OperatingSystemVersion = '26.2'                 ; Description = 'Android tablet – setlists (service account)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.62'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXASWICLY001' ; SamAccountName = 'EXASWICLY001$' ; Type = 'Switch'        ; Role = 'SWI' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASWICLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'Cisco Catalyst 9300'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'TPLink 48port managed switch; admin:tplink1987' ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
      @{ Name = 'EXARACCLY001' ; SamAccountName = 'EXARACCLY001$' ; Type = 'RAC'           ; Role = 'RAC' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXARACCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','BMC')               ; OS = 'HP iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell DRAC for EXADCCLY001; dracadmin:WetWet87'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Miami Infra
      @{ Name = 'EXACOFMIA001' ; SamAccountName = 'EXACOFMIA001$' ; Type='Embedded'        ; Role='VND'   ; Site = 'MIA' ; Location = 'Miami, FL'                    ; DNSHostName = 'EXACOFMIA001.example.net' ; OU = @('Locations','US','Florida','Miami','Vending Machhines')      ; OS = 'VxWorks'                          ; OperatingSystemVersion = '7.25.09'              ; Description = 'Networked Cuban Covfefe machine svc_coffee'     ; LastLogon = 'N/A'                      ; IPAddress='10.30.10.50'       ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

      ## Brockville, Ontario Infra
      @{ Name ='EXADONBRK001'  ; SamAccountName = 'EXADONBRK001$' ; Type='Embedded'        ; Role='VND'   ; Site = 'BRK' ; Location = 'Brockville, ON'               ; DNSHostName = 'EXADONBRK001.example.net' ; OU = @('Locations','CA','Ontario','Brockville','Vending Machhines') ; OS = 'VxWorks'                          ; OperatingSystemVersion = '7.25.09'              ; Description = 'Donut vending machine (Tim Hortons compatible)' ; LastLogon = 'N/A'                      ; IPAddress='10.20.10.50'       ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }
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
    Set-StatusBar "Enumerating DCs..." -spinner
    $dcs = Invoke-AD { Get-ADDomainController -Filter * -Server $domain -ErrorAction Stop } -SuppressError

    ## Get Users
    Set-StatusBar "Enumerating Users..." -spinner
    $users = Invoke-AD { Get-ADUser -Filter * -Server $domain -Properties DisplayName,EmailAddress,Title,Department,Enabled,LockedOut,DistinguishedName -ErrorAction Stop } -SuppressError

    ## Get Groups
    Set-StatusBar "Enumerating Groups..." -spinner
    $groups = Invoke-AD { Get-ADGroup -Filter * -Server $domain -Properties Description,GroupCategory,GroupScope,Members,DistinguishedName -ErrorAction Stop } -SuppressError

    ## Get Computers
    Set-StatusBar "Enumerating Computers..." -spinner
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
    Set-StatusBar "Setting AD data..." -spinner
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
  function New-FakeSid { $rid = Get-Random -Minimum 1000 -Maximum 65535 ; "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999), (Get-Random -Max 999999999), (Get-Random -Max 999999999), $rid }

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

   ## Add this RIGHT BEFORE the DC section in Build-DomainContent:
   Debug-Log ": About to build DCs - rawDCs exists: $($null -ne $Script:rawDCs), count: $($Script:rawDCs.Count), domain: $domain" -Type "Info"

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

    $tn = [Terminal.Gui.Trees.TreeNode]::new($text)
    $tn.Tag = $tag

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
  if ($Script:FilterOptions.ShowGroups) {
    $domainGroups = $Script:Groups | Where-Object { $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain }
    if ($domainGroups.Count -gt 0) {
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

        # Add group members
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

        if ($groupNode.Children.Count -gt 0) { $groupsNode.Children.Add($groupNode) }
      }
    }
  }

  # --- Domain Controllers ---
  if ($Script:FilterOptions.ShowDCs -and $Script:rawDCs -and $Script:rawDCs.Count -gt 0) {
    ## Filter DCs for current domain
    $dcsInDomain = $Script:rawDCs | Where-Object { $_.Domain -eq $domain }

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

      ## Initialize children collection
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
  if ($Script:FilterOptions.ShowComputers) {
    $domainComputers = $Script:Computers | Where-Object { $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain }
    if ($domainComputers.Count -gt 0) {
      $computerNode = Get-OrCreateChildNode -Parent $domainNode -Name "Computers" -FullPath "$domain/_Computers" -NodeType 'container'
      $byType = $domainComputers | Group-Object ComputerType
      foreach ($typeGroup in $byType) {
        $typeNode = [Terminal.Gui.Trees.TreeNode]::new("$($typeGroup.Name)s")
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
        if ($typeNode.Children.Count -gt 0) { $computerNode.Children.Add($typeNode) }
      }
    }
  }

  Debug-Log ": Finished building content for domain $domain" -Type "Success"
}

#### This is never called - is it dead code?

function New-TreeObjectNode {
    param(
        [string]$Type,   # user | group | computer | dc | ou
        [object]$Object  # the actual object
    )

    $icon = switch ($Type) {
        'user'     { $Script:Icons.User }
        'group'    { $Script:Icons.Group }
        'computer' { $Script:Icons.Computer }
        'dc'       { $Script:Icons.DC }
        'ou'       { $Script:Icons.OU }
        default    { '' }
    }

    $node = [Terminal.Gui.Trees.TreeNode]::new("$icon $($Object.Name)")
    $node.Tag = @{
        Type   = $Type
        Object = $Object
    }

    return $node
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
    # Create standalone label
    $lblStatus = Manage-FilterStatusLabel -Action 'Create'

    # Create label inside panel
    $lblStatus = Manage-FilterStatusLabel -Action 'Create' -X 1 -Y 15 -InPanel

    # Update existing label
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
        [object]$Width = 40,  # Can be int or Dim object

        [switch]$InPanel
    )

    if ($Action -eq 'Create') {
        Debug-Log ": Creating filter status label" -Type "Info"

        try {
            $lblStatus = [Terminal.Gui.Label]::new("")
            $lblStatus.X = $X
            $lblStatus.Y = $Y

            # Handle Width - can be int or Dim object
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

        # Build active filters list
        $activeFilters = @()

        if (-not $Script:FilterOptions.ShowEnabledUsers) {
            $activeFilters += "No Enabled"
        }
        if (-not $Script:FilterOptions.ShowDisabledUsers) {
            $activeFilters += "No Disabled"
        }
        if (-not $Script:FilterOptions.ShowGroups) {
            $activeFilters += "No Groups"
        }
        if (-not $Script:FilterOptions.ShowDCs) {
            $activeFilters += "No DCs"
        }
        if (-not $Script:FilterOptions.ShowComputers) {
            $activeFilters += "No Computers"
        }
        if ($Script:FilterOptions.NameFilter) {
            $activeFilters += "Name:$($Script:FilterOptions.NameFilter)"
        }

        # Update label text
        if ($activeFilters.Count -gt 0) {
            $Label.Text = "Active Filters: " + ($activeFilters -join ", ")
        } else {
            $Label.Text = "No filters active (showing all)"
        }

        Debug-Log ": Filter status label updated: $($activeFilters.Count) filters active" -Type "Info"
    }
}

## ==================== LDAP FILTER HELPER ====================
function Get-LDAPFilteredUsers {
  <#
  .SYNOPSIS
  Apply LDAP-style filters to user collection

  .PARAMETER FilterType
  Type of filter to apply

  .PARAMETER Users
  User collection to filter

  .EXAMPLE
  Get-LDAPFilteredUsers -FilterType 'PasswordExpired' -Users $Script:Users
  #>
  param(
    [Parameter(Mandatory=$true)]
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

    [array]$Users = $Script:Users
  )

  $now = Get-Date

  switch ($FilterType) {
    'All' {
      return $Users
    }

    'LockedOnly' {
      return $Users | Where-Object { $_.LockedOut -or $_.Locked }
    }

    'DisabledOnly' {
      return $Users | Where-Object { $_.Disabled -or -not $_.Enabled }
    }

    'EnabledOnly' {
      return $Users | Where-Object { -not $_.Disabled -and $_.Enabled }
    }

    'NeverLoggedIn' {
      if ($Script:DemoMode) {
        # Demo: users with no LastLogon or very old LastLogon
        return $Users | Where-Object {
          -not $_.LastLogonDate -or
          ($_.LastLogonDate -lt (Get-Date).AddYears(-5))
        }
      } else {
        return $Users | Where-Object {
          -not $_.LastLogonDate -or
          $_.LastLogonDate -eq [DateTime]::MinValue
        }
      }
    }

    'NoManager' {
      return $Users | Where-Object {
        [string]::IsNullOrWhiteSpace($_.Manager)
      }
    }

    'PasswordExpired' {
      if ($Script:DemoMode) {
        # Demo: users with MustChangePassword or pwdLastSet = 0
        return $Users | Where-Object {
          $_.MustChangePassword -or
          $_.PasswordExpired -or
          $_.pwdLastSet -eq 0
        }
      } else {
        return $Users | Where-Object { $_.PasswordExpired -eq $true }
      }
    }

    'PasswordExpiring72h' {
      # Calculate password expiration (typically 90 days from last set)
      $expiryDays = 90  # Default policy
      $warningWindow = $now.AddHours(72)

      return $Users | Where-Object {
        if ($_.PasswordNeverExpires) { return $false }
        if (-not $_.PasswordLastSet) { return $false }

        $expiryDate = $_.PasswordLastSet.AddDays($expiryDays)
        $expiryDate -le $warningWindow -and $expiryDate -gt $now
      }
    }

    'PasswordNeverExpires' {
      return $Users | Where-Object { $_.PasswordNeverExpires -eq $true }
    }

    'AccountExpired' {
      return $Users | Where-Object {
        $_.AccountExpirationDate -and
        $_.AccountExpirationDate -lt $now
      }
    }

    'AccountExpiring30d' {
      $window = $now.AddDays(30)
      return $Users | Where-Object {
        $_.AccountExpirationDate -and
        $_.AccountExpirationDate -gt $now -and
        $_.AccountExpirationDate -le $window
      }
    }

    'StaleAccounts90d' {
      $staleDate = $now.AddDays(-90)
      return $Users | Where-Object {
        $_.LastLogonDate -and
        $_.LastLogonDate -lt $staleDate
      }
    }

    'EmptyEmail' {
      return $Users | Where-Object {
        [string]::IsNullOrWhiteSpace($_.EmailAddress) -and
        [string]::IsNullOrWhiteSpace($_.mail)
      }
    }

    'EmptyDepartment' {
      return $Users | Where-Object {
        [string]::IsNullOrWhiteSpace($_.Department)
      }
    }
  }
}

### Never called, but surely can be mergeed with sister user function above...?
function Get-LDAPFilteredGroups {
  <#
  .SYNOPSIS
  Apply LDAP-style filters to group collection
  #>
  param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('All', 'EmptyGroups', 'SecurityGroups', 'DistributionGroups')]
    [string]$FilterType,

    [array]$Groups = $Script:Groups
  )

  switch ($FilterType) {
    'All' {
      return $Groups
    }

    'EmptyGroups' {
      if ($Script:DemoMode) {
        return $Groups | Where-Object {
          $groupName = $_.Name
          $memberCount = ($Script:Users | Where-Object {
            $_.Groups -contains $groupName -or
            $_.MemberOf -contains $groupName
          }).Count
          $memberCount -eq 0
        }
      } else {
        return $Groups | Where-Object {
          -not $_.Members -or
          $_.Members.Count -eq 0
        }
      }
    }

    'SecurityGroups' {
      return $Groups | Where-Object {
        $_.GroupCategory -eq 'Security' -or
        $_.Type -eq 'Security'
      }
    }

    'DistributionGroups' {
      return $Groups | Where-Object {
        $_.GroupCategory -eq 'Distribution' -or
        $_.Type -eq 'Distribution'
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
    $filtered = Get-LDAPFilteredUsers -FilterType $Script:FilterOptions.QuickFilter -Users $filtered
  }

  ## Apply enabled/disabled checkboxes (if no quick filter overriding)
  if (-not $Script:FilterOptions.QuickFilter -or $Script:FilterOptions.QuickFilter -eq 'All') {
    $filtered = $filtered | Where-Object {
      ($_.Disabled -and $Script:FilterOptions.ShowDisabledUsers) -or
      (-not $_.Disabled -and $Script:FilterOptions.ShowEnabledUsers)
    }
  }

  ## Apply name filter with operator
  $nameFilter = $Script:FilterOptions.NameFilter.Trim()
  if ($nameFilter) {
    $operator = $Script:FilterOptions.NameOperator

    switch ($operator) {
      'Contains' {
        $filtered = $filtered | Where-Object {
          $_.Name -like "*$nameFilter*" -or
          $_.EmailAddress -like "*$nameFilter*" -or
          $_.Title -like "*$nameFilter*"
        }
      }
      'StartsWith' {
        $filtered = $filtered | Where-Object {
          $_.Name -like "$nameFilter*" -or
          $_.EmailAddress -like "$nameFilter*" -or
          $_.Title -like "$nameFilter*"
        }
      }
      'EndsWith' {
        $filtered = $filtered | Where-Object {
          $_.Name -like "*$nameFilter" -or
          $_.EmailAddress -like "*$nameFilter" -or
          $_.Title -like "*$nameFilter"
        }
      }
      'Equals' {
        $filtered = $filtered | Where-Object {
          $_.Name -eq $nameFilter -or
          $_.EmailAddress -eq $nameFilter -or
          $_.SamAccountName -eq $nameFilter
        }
      }
    }
  }

  return $filtered
}

# ------------------------- Filter Panel (Add to main window)-------------------------
function Create-FilterPanel {

  ## Initialize FilterOptions with new fields
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
      NameOperator      = "Contains"  # NEW
      QuickFilter       = "All"       # NEW
      SortBy            = "Name"
      SortDescending    = $false
    }
  }

  ## Create frame
  $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
  $filterFrame.X = 32
  $filterFrame.Y = 1
  $filterFrame.Width = 40
  $filterFrame.Height = 27  # Increased for separator lines and better spacing
  $y = 0

  ## ==================== Name Filter with Operator ====================
  $lblNameFilter = [Terminal.Gui.Label]::new("Search Name/Email/Title:")
  $lblNameFilter.X=1; $lblNameFilter.Y=$y
  $filterFrame.Add($lblNameFilter)
  $y+=1

  ## Operator dropdown
  $lblOperator = [Terminal.Gui.Label]::new("Match:")
  $lblOperator.X=1; $lblOperator.Y=$y
  $filterFrame.Add($lblOperator)

  $cmbOperator = [Terminal.Gui.ComboBox]::new()
  $cmbOperator.X=8; $cmbOperator.Y=$y; $cmbOperator.Width=13
  $cmbOperator.SetSource([string[]]@("Contains", "StartsWith", "EndsWith", "Equals"))
  $cmbOperator.Text = [NStack.ustring]::Make($Script:FilterOptions.NameOperator)
  $cmbOperator.add_SelectedItemChanged({
    $Script:FilterOptions.NameOperator = $cmbOperator.Text.ToString()
  })
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
  $chkEnabled = [Terminal.Gui.CheckBox]::new("Show Enabled Users")
  $chkEnabled.X=1; $chkEnabled.Y=$y; $chkEnabled.Checked=$Script:FilterOptions.ShowEnabledUsers
  $chkEnabled.add_Toggled({ $Script:FilterOptions.ShowEnabledUsers = $chkEnabled.Checked })
  $filterFrame.Add($chkEnabled)
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
  $y+=1

  ## Separator line (visual spacing)
  $lblSep3 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep3.X=1; $lblSep3.Y=$y
  $filterFrame.Add($lblSep3)
  $y+=1

  ## ==================== Apply/Reset Buttons ====================
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
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  })
  $filterFrame.Add($btnApplyFilter)

  $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
  $btnResetFilter.X=21; $btnResetFilter.Y=$y
  $btnResetFilter.add_Clicked({
    Debug-Log "Resetting filters..." -Type "Info"

    ## Reset all filters
    $Script:FilterOptions.ShowDisabledUsers = $true
    $Script:FilterOptions.ShowEnabledUsers = $true
    $Script:FilterOptions.ShowLockedUsers = $true
    $Script:FilterOptions.ShowGroups = $true
    $Script:FilterOptions.ShowDCs = $true
    $Script:FilterOptions.ShowComputers = $true
    $Script:FilterOptions.ShowOUs = $true
    $Script:FilterOptions.NameFilter = ""
    $Script:FilterOptions.NameOperator = "Contains"
    $Script:FilterOptions.QuickFilter = "All"
    $Script:FilterOptions.SortBy = "Name"
    $Script:FilterOptions.SortDescending = $false

    ## Reset UI controls
    $chkEnabled.Checked = $true
    $chkDisabled.Checked = $true
    $chkGroups.Checked = $true
    $chkDCs.Checked = $true
    $txtNameFilter.Text = ""
    $cmbOperator.SelectedItem = 0
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
  })
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
  })
  $dlg.AddButton($btnApply)

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)

  [Terminal.Gui.Application]::Run($dlg)
}

##bums

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
    IsLinux = $false
    IsMacOS = $false
    OSName = "Unknown"
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
  $hasADModule = $null -ne (Get-Module -ListAvailable -Name ActiveDirectory)
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

  ## Tab definitions
  $tabs = @(
    @{ Name = "System Info"; Generator = { Get-SystemInfoText } }
    @{ Name = "Domain Controllers"; Generator = { Get-DCStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Replication"; Generator = { Get-ReplicationStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "DNS Records"; Generator = { Get-DNSStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "SYSVOL/NETLOGON"; Generator = { Get-SysvolStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "FSMO Roles"; Generator = { Get-FSMOStatusText -Domain $Script:ADHealthDomain } }
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
  })
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
  })
  $dialog.Add($btnExport)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::AnchorEnd(10)
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnClose.add_Clicked({
    Debug-Log ": AD Health dialog closed" -Type "Info"
    [Terminal.Gui.Application]::RequestStop()
  })
  $dialog.Add($btnClose)

  ## Run dialog
  Debug-Log ": Running AD Health dialog" -Type "Info"
  [Terminal.Gui.Application]::Run($dialog)
}


## ==================== CONTENT GENERATORS ====================

function Get-SystemInfoText {
  $output = @()
  $output += "SYSTEM INFORMATION"
  $output += "=" * 80
  $output += ""

  ## Operating System
  $output += "Operating System:"
  $output += "  Platform:  $($Script:ADHealthOS.OSName)"
  $output += "  Windows:   $($Script:ADHealthOS.IsWindows)"
  $output += "  Linux:     $($Script:ADHealthOS.IsLinux)"
  $output += "  macOS:     $($Script:ADHealthOS.IsMacOS)"
  $output += ""

  ## PowerShell Version
  $output += "PowerShell:"
  $output += "  Version:   $($PSVersionTable.PSVersion)"
  $output += "  Edition:   $($PSVersionTable.PSEdition)"
  $output += ""

  ## Module Availability
  $output += "PowerShell Modules:"
  $hasAD = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
  $output += "  ActiveDirectory:    $(if ($hasAD) { '✓ Available' } else { '✗ Not Found' })"

  $hasTerminalIcons = Get-Module -ListAvailable -Name Terminal-Icons -ErrorAction SilentlyContinue
  $output += "  Terminal-Icons:     $(if ($hasTerminalIcons) { '✓ Available' } else { '✗ Not Found' })"

  $hasNerdFonts = Get-Module -ListAvailable -Name NerdFonts -ErrorAction SilentlyContinue
  $output += "  NerdFonts:          $(if ($hasNerdFonts) { '✓ Available' } else { '✗ Not Found' })"
  $output += ""

  ## Tool Availability
  $output += "External Tools:"
  foreach ($tool in $Script:ADHealthTools.Keys | Sort-Object) {
    $status = if ($Script:ADHealthTools[$tool]) { "✓ Available" } else { "✗ Not Found" }
    $output += "  {0,-20} {1}" -f "${tool}:", $status
  }
  $output += ""

  ## Demo Mode
  $output += "Application Mode:"
  $output += "  Demo Mode:  $($Script:DemoMode)"
  $output += "  Domain:     $($Script:ADHealthDomain)"
  $output += ""

  ## Data Summary (from Script variables)
  if ($Script:DemoMode) {
    $output += "Demo Data Summary:"
    $output += "  Users:      $(if ($Script:rawUsers) { $Script:rawUsers.Count } else { 0 })"
    $output += "  DCs:        $(if ($Script:rawDCs) { $Script:rawDCs.Count } elseif ($Script:DCs) { $Script:DCs.Count } else { 0 })"
    $output += "  Computers:  $(if ($Script:rawComputers) { $Script:rawComputers.Count } else { 0 })"
    $output += "  Groups:     $(if ($Script:rawDemoGroups) { $Script:rawDemoGroups.Count } else { 0 })"
    $output += ""
  }

  ## Installation Instructions
  if (-not $hasAD) {
    $output += "⚠ ActiveDirectory Module Not Found"
    $output += ""
    if ($Script:ADHealthOS.IsWindows) {
      $output += "To install on Windows:"
      $output += "  1. Open PowerShell as Administrator"
      $output += "  2. Run: Add-WindowsFeature RSAT-AD-PowerShell"
      $output += "     OR: Install-WindowsFeature RSAT-AD-PowerShell"
    } elseif ($Script:ADHealthOS.IsLinux) {
      $output += "To install on Linux:"
      $output += "  1. Install realmd: sudo apt install realmd sssd sssd-tools adcli"
      $output += "  2. Join domain: sudo realm join domain.com"
    } elseif ($Script:ADHealthOS.IsMacOS) {
      $output += "To install on macOS:"
      $output += "  1. Install via Homebrew: brew install realmd"
      $output += "  2. Configure AD integration"
    }
    $output += ""
  }

  return ($output -join "`n")
}


function Get-DCStatusText {
  param([string]$Domain)

  $output = @()
  $output += "DOMAIN CONTROLLER STATUS"
  $output += "=" * 80
  $output += ""

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
      $output += "  Total DCs:     $totalCount"
      $output += "  Online:        $onlineCount"
      $output += "  Offline:       $($totalCount - $onlineCount)"
      $output += "  Health:        $($result.Health)"
      $output += ""
      $output += "Ports: LDAP (389), Kerberos (88), SMB (445), DNS (53)"
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
  $output += "REPLICATION STATUS"
  $output += "=" * 80
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
  $output += "DNS RECORDS STATUS"
  $output += "=" * 80
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
  $output += "SYSVOL/NETLOGON STATUS"
  $output += "=" * 80
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


function Get-FSMOStatusText {
  param([string]$Domain)

  $output = @()
  $output += "FSMO ROLES"
  $output += "=" * 80
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


function Get-GPOStatusText {
  param([string]$Domain)

  $output = @()
  $output += "GROUP POLICY STATUS"
  $output += "=" * 80
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
  $tools = @{
    'repadmin.exe' = $false
    'dfsrdiag.exe' = $false
    'dcdiag.exe'   = $false
    'nslookup.exe' = $false
    'ping.exe'     = $false
    'netstat.exe'  = $false
    'csvde.exe'    = $false
    'portqry.exe'  = $false
    'nltest.exe'   = $false
  }

  foreach ($tool in $tools.Keys) {
    try {
      ## nslookup is built-in on Windows, won't show up in Get-Command
      if ($tool -eq 'nslookup.exe' -and $Script:ADHealthOS.IsWindows) {
        $tools[$tool] = $true
      } else {
        $null = Get-Command $tool -ErrorAction Stop
        $tools[$tool] = $true
      }
    } catch {
      $tools[$tool] = $false
    }
  }

  return $tools
}


function Invoke-ExternalCommand {
  param([string]$Exe, [string]$Args)

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = $Args
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    $out = $proc.StandardOutput.ReadToEnd()
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    return @{ ExitCode = $proc.ExitCode; StdOut = $out; StdErr = $err }
  } catch {
    return @{ ExitCode = -1; StdOut = ""; StdErr = $_.Exception.Message }
  }
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
        $smbOK = (Test-NetConnection -ComputerName $name -Port 445 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        $dnsOK = (Test-NetConnection -ComputerName $name -Port 53 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
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
    $r = Invoke-ExternalCommand -Exe "repadmin.exe" -Args "/replsummary"

    if ($r.ExitCode -eq 0) {
      $lines = ($r.StdOut -split "`r?`n") | Where-Object { $_ -match '\S' } | Select-Object -First 20
      $summary = $lines
      $details = $r.StdOut

      if ($r.StdOut -match "error|fail") {
        $health = "FAIL"
      }
    } else {
      $summary += "repadmin returned error code $($r.ExitCode)"
      $details = $r.StdErr
      $health = "WARN"
    }
  } else {
    ## Fallback to AD cmdlet
    try {
      $failures = Get-ADReplicationFailure -Scope Domain -Target $Domain -ErrorAction Stop

      if ($failures) {
        foreach ($f in $failures) {
          $summary += "FAIL: $($f.Server) - $($f.FirstFailureMessage)"
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
    $gpos = Get-GPO -All -Domain $Domain -ErrorAction Stop

    $summary += "Total GPOs: $($gpos.Count)"
    $details = "GPO List (first 50):`n"

    foreach ($gpo in $gpos | Select-Object -First 50) {
      $summary += "$($gpo.DisplayName)"
      $details += "  $($gpo.DisplayName) (ID: $($gpo.Id))`n"
    }

    if ($gpos.Count -gt 50) {
      $summary += "... and $($gpos.Count - 50) more"
      $details += "`n... and $($gpos.Count - 50) more GPOs"
    }

  } catch {
    $summary += "Error checking GPOs: $($_.Exception.Message)"
    $details = "Error: $($_.Exception.Message)"
    $health = "FAIL"
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
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
    })

    ## --- Close ---
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

  [Terminal.Gui.Application]::Run($dlg)
  return $Script:actualPassword
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
  Debug-Log ": DemoMode = $($Script:DemoMode)" -Type "Info"

  ## -------------------------{ Safety: AD availability }-------------------------
  if (-not $Script:DemoMode) {
    Debug-Log ": Production mode - checking AD module" -Type "Info"
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
      Debug-Log ": ActiveDirectory module not available" -Type "Error"
      Show-Modal "LAPS Error" "ActiveDirectory module not available."
      return
    }

    try {
      Import-Module ActiveDirectory -ErrorAction Stop
      Debug-Log ": ActiveDirectory module imported" -Type "Success"
    } catch {
      Debug-Log ": Failed to import AD module: $($_.Exception.Message)" -Type "Error"
      Show-Modal "LAPS Error" "Failed to import ActiveDirectory module:`n`n$($_.Exception.Message)"
      return
    }
  }

  ## -------------------------{ Modal }-------------------------
  Debug-Log ": Creating LAPS dialog" -Type "Info"

  $dialog = [Terminal.Gui.Dialog]::new("LAPS Password Lookup", 50, 25)

  $lblSearch = [Terminal.Gui.Label]::new("Computer name (blank = all):")
  $lblSearch.X = 1
  $lblSearch.Y = 1
  $dialog.Add($lblSearch)

  $txtSearch = [Terminal.Gui.TextField]::new("")
  $txtSearch.X = 1
  $txtSearch.Y = 2
  $txtSearch.Width = 30
  $dialog.Add($txtSearch)

  $lstComputers = [Terminal.Gui.ListView]::new()
  $lstComputers.X = 1
  $lstComputers.Y = 4
  $lstComputers.Width = 30
  $lstComputers.Height = 16
  $dialog.Add($lstComputers)

  ## Script-scoped variable to hold computers
  $Script:LAPSComputers = @()

  ## ---------------- Data loader ----------------
  $loadComputers = {
    param($filter)

    Debug-Log ": Loading LAPS computers with filter: '$filter'" -Type "Info"

    try {
      if ($Script:DemoMode) {
        Debug-Log ": Demo mode - loading from rawComputers" -Type "Info"

        ## Filter computers that have LAPS properties
        $allComputers = $Script:rawComputers | Where-Object {
          $null -ne $_.'msLAPS-Password' -and
          $_.'msLAPS-Password' -ne '' -and
          $_.'msLAPS-Password' -ne $null
        }

        Debug-Log ": Found $($allComputers.Count) computers with LAPS in rawComputers" -Type "Info"

        ## Apply name filter if specified
        if (-not [string]::IsNullOrWhiteSpace($filter)) {
          $Script:LAPSComputers = $allComputers | Where-Object { $_.Name -like "*$filter*" }
          Debug-Log ": Filtered to $($Script:LAPSComputers.Count) computers matching '$filter'" -Type "Info"
        } else {
          $Script:LAPSComputers = $allComputers
          Debug-Log ": No filter - showing all $($Script:LAPSComputers.Count) LAPS computers" -Type "Info"
        }

        ## Convert hashtables to PSCustomObjects for consistent access
        $Script:LAPSComputers = $Script:LAPSComputers | ForEach-Object {
          if ($_ -is [hashtable]) {
            [PSCustomObject]$_
          } else {
            $_
          }
        }

      } else {
        ## Production mode
        Debug-Log ": Production mode - querying AD" -Type "Info"

        if ([string]::IsNullOrWhiteSpace($filter)) {
          $Script:LAPSComputers = Get-ADComputer -Filter * -Properties 'ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime' |
            Where-Object { $_.'ms-Mcs-AdmPwd' }
        } else {
          $Script:LAPSComputers = Get-ADComputer -Filter "Name -like '*$filter*'" -Properties 'ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime' |
            Where-Object { $_.'ms-Mcs-AdmPwd' }
        }

        Debug-Log ": Found $($Script:LAPSComputers.Count) computers with LAPS in AD" -Type "Info"
      }

      ## Update ListView
      if ($Script:LAPSComputers.Count -eq 0) {
        Debug-Log ": No LAPS computers found" -Type "Warn"
        $lstComputers.SetSource(@("(No computers with LAPS found)"))
      } else {
        $computerNames = $Script:LAPSComputers | ForEach-Object { $_.Name }
        Debug-Log ": Setting ListView with $($computerNames.Count) computer names" -Type "Info"
        $lstComputers.SetSource($computerNames)
      }

    } catch {
      $errMsg = $_.Exception.Message
      Debug-Log ": Error loading LAPS computers: $errMsg" -Type "Error"
      Debug-Log ": Stack trace: $($_.ScriptStackTrace)" -Type "Error"
      Show-Modal "Error" "Failed to load LAPS computers:`n`n$errMsg"
    }
  }

  ## Initial load
  Debug-Log ": Performing initial computer load" -Type "Info"
  & $loadComputers ""

  ## ---------------- Buttons ----------------
  $btnSearch = [Terminal.Gui.Button]::new("Search")
  $btnSearch.X = 32
  $btnSearch.Y = 2
  $btnSearch.add_Clicked({
    $searchText = $txtSearch.Text.ToString()
    Debug-Log ": Search button clicked with text: '$searchText'" -Type "Info"
    & $loadComputers $searchText
  })
  $dialog.Add($btnSearch)

  $btnView = [Terminal.Gui.Button]::new("View")
  $btnView.X = 10
  $btnView.Y = 21
  $dialog.Add($btnView)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = 30
  $btnClose.Y = 21
  $btnClose.add_Clicked({
    Debug-Log ": LAPS dialog closed" -Type "Info"
    [Terminal.Gui.Application]::RequestStop()
  })
  $dialog.Add($btnClose)

  ## ---------------- View handler ----------------
  $btnView.add_Clicked({
    Debug-Log ": View button clicked" -Type "Info"

    if ($lstComputers.SelectedItem -lt 0) {
      Debug-Log ": No computer selected" -Type "Warn"
      Show-Modal "No Selection" "Please select a computer from the list."
      return
    }

    if ($Script:LAPSComputers.Count -eq 0) {
      Debug-Log ": No computers available to view" -Type "Warn"
      Show-Modal "No Data" "No computers with LAPS available."
      return
    }

    try {
      $computer = $Script:LAPSComputers[$lstComputers.SelectedItem]
      Debug-Log ": Selected computer: $($computer.Name)" -Type "Info"

      ## Get LAPS properties (handle both naming conventions)
      if ($Script:DemoMode) {
        $lapsUser = $computer.'msLAPS-AccountName'
        $lapsPassword = $computer.'msLAPS-Password'
        $lapsExpiry = $computer.'msLAPS-PasswordExpirationTime'
      } else {
        $lapsUser = if ($computer.'ms-Mcs-AdmPwd') { 'Administrator' } else { 'Unknown' }
        $lapsPassword = $computer.'ms-Mcs-AdmPwd'
        $lapsExpiry = $computer.'ms-Mcs-AdmPwdExpirationTime'
      }

      Debug-Log ": LAPS User: $lapsUser" -Type "Info"
      Debug-Log ": LAPS Password length: $($lapsPassword.Length)" -Type "Info"
      Debug-Log ": LAPS Expiry (raw): $lapsExpiry" -Type "Info"

      ## Calculate expiry
      if ($lapsExpiry) {
        try {
          $expires = [DateTime]::FromFileTimeUtc([long]$lapsExpiry)
          $daysLeft = [Math]::Round(($expires - (Get-Date)).TotalDays)
          Debug-Log ": Password expires: $expires ($daysLeft days)" -Type "Info"
        } catch {
          Debug-Log ": Failed to parse expiry: $($_.Exception.Message)" -Type "Warn"
          $expires = Get-Date
          $daysLeft = 999
        }
      } else {
        $expires = Get-Date
        $daysLeft = 999
      }

      ## Expiry warning
      $warn = if ($daysLeft -le 3) {
        "⚠ WARNING: Password expires in $daysLeft day(s)"
      } elseif ($daysLeft -le 7) {
        "⏰ Password expires in $daysLeft days"
      } else {
        "Expires in $daysLeft days"
      }

      ## Create detail dialog (more compact)
      $detailDlg = [Terminal.Gui.Dialog]::new("LAPS Details - $($computer.Name)", 50, 14)

      ## Computer name
      $lblComputer = [Terminal.Gui.Label]::new("Computer: $($computer.Name)")
      $lblComputer.X = 2
      $lblComputer.Y = 1
      $detailDlg.Add($lblComputer)

      ## LAPS User (inline with aligned spacing)
      $lblUserLabel = [Terminal.Gui.Label]::new("LAPS User:")
      $lblUserLabel.X = 2
      $lblUserLabel.Y = 3
      $detailDlg.Add($lblUserLabel)

      $lblUserValue = [Terminal.Gui.Label]::new($lapsUser)
      $lblUserValue.X = 18  # Aligned at column 18
      $lblUserValue.Y = 3
      $lblUserValue.Width = 25
      $detailDlg.Add($lblUserValue)

      $btnCopyUser = [Terminal.Gui.Button]::new("📋")
      $btnCopyUser.X = 40  ## button placement horiztonally
      $btnCopyUser.Y = 3
      $btnCopyUser.Width = 4
      $btnCopyUser.add_Clicked({
        Set-Clipboard -Value $lapsUser
        Debug-Log ": LAPS user copied to clipboard" -Type "Info"
        Show-Modal "Copied" "Username copied to clipboard."
      })
      $detailDlg.Add($btnCopyUser)

      ## LAPS Password (inline with reveal/hide) - ALIGNED
      $lblPasswordLabel = [Terminal.Gui.Label]::new("LAPS Password:")
      $lblPasswordLabel.X = 2
      $lblPasswordLabel.Y = 5
      $detailDlg.Add($lblPasswordLabel)

      ## Create masked version
      $masked = ('*' * $lapsPassword.Length)

      ## Use script-scoped variable for toggle state (IMPORTANT for closure)
      $Script:LAPSPasswordRevealed = $false

      $lblPasswordValue = [Terminal.Gui.Label]::new($masked)
      $lblPasswordValue.X = 18  # Aligned at column 18
      $lblPasswordValue.Y = 5
      $lblPasswordValue.Width = 22
      $detailDlg.Add($lblPasswordValue)

      $btnReveal = [Terminal.Gui.Button]::new("👁")
      $btnReveal.X = 34
      $btnReveal.Y = 5
      $btnReveal.Width = 3
      $btnReveal.add_Clicked({
        if ($Script:LAPSPasswordRevealed) {
          ## Hide password
          $lblPasswordValue.Text = [NStack.ustring]::Make($masked)
          $btnReveal.Text = [NStack.ustring]::Make("👁")
          $Script:LAPSPasswordRevealed = $false
          Debug-Log ": Password hidden" -Type "Info"
        } else {
          ## Show password
          $lblPasswordValue.Text = [NStack.ustring]::Make($lapsPassword)
          $btnReveal.Text = [NStack.ustring]::Make("🙈")
          $Script:LAPSPasswordRevealed = $true
          Debug-Log ": Password revealed" -Type "Info"
        }
      })
      $detailDlg.Add($btnReveal)

      $btnCopyPassword = [Terminal.Gui.Button]::new("📋")
      $btnCopyPassword.X = 40  ## Horizontal alignment
      $btnCopyPassword.Y = 5
      $btnCopyPassword.Width = 4
      $btnCopyPassword.add_Clicked({
        Set-Clipboard -Value $lapsPassword
        Debug-Log ": LAPS password copied to clipboard" -Type "Info"
        Show-Modal "Copied" "Password copied to clipboard."
      })
      $detailDlg.Add($btnCopyPassword)

      ## Expiry warning
      $lblExpiry = [Terminal.Gui.Label]::new($warn)
      $lblExpiry.X = 2
      $lblExpiry.Y = 7
      $lblExpiry.Width = 62
      $detailDlg.Add($lblExpiry)

      ## Force Rotation button - ALWAYS show (works in both demo and production)
      $btnRotate = [Terminal.Gui.Button]::new("🔄 Force Rotation")
      $btnRotate.X = 2
      $btnRotate.Y = 9
      $btnRotate.add_Clicked({
        if ($Script:DemoMode) {
          ## Demo mode - just show what would happen
          Debug-Log ": LAPS rotation simulated (demo mode) for $($computer.Name)" -Type "Info"
          Show-Modal "Demo Mode" "In production, this would:`n`n1. Set ms-Mcs-AdmPwdExpirationTime to 0`n2. Force password rotation on next GP refresh`n`nRun 'gpupdate /force' on the computer to trigger immediately."
        } else {
          ## Production mode - actually do it
          try {
            Set-ADComputer -Identity $computer.Name -Replace @{'ms-Mcs-AdmPwdExpirationTime' = '0'}
            Debug-Log ": LAPS password rotation forced for $($computer.Name)" -Type "Success"
            Show-Modal "Rotation Requested" "LAPS password will rotate on next Group Policy refresh.`n`nRun 'gpupdate /force' on $($computer.Name) to trigger immediately.`n`nRemote option:`nInvoke-GPUpdate -Computer $($computer.Name) -Force"
          } catch {
            Debug-Log ": Failed to force LAPS rotation: $($_.Exception.Message)" -Type "Error"
            Show-Modal "Error" "Failed to force rotation:`n`n$($_.Exception.Message)"
          }
        }
      })
      $detailDlg.Add($btnRotate)

      ## Close button
      $btnOk = [Terminal.Gui.Button]::new("Close")
      $btnOk.X = 26
      $btnOk.Y = 9
      $btnOk.add_Clicked({
        Debug-Log ": LAPS detail dialog closed" -Type "Info"
        [Terminal.Gui.Application]::RequestStop()
      })
      $detailDlg.Add($btnOk)

      Debug-Log ": Running LAPS detail dialog" -Type "Info"
      [Terminal.Gui.Application]::Run($detailDlg)

    } catch {
      $errMsg = $_.Exception.Message
      Debug-Log ": Error showing LAPS details: $errMsg" -Type "Error"
      Debug-Log ": Stack trace: $($_.ScriptStackTrace)" -Type "Error"
      Show-Modal "Error" "Failed to show LAPS details:`n`n$errMsg"
    }
  })

  Debug-Log ": Running LAPS search dialog" -Type "Info"
  [Terminal.Gui.Application]::Run($dialog)
  Debug-Log ": LAPS search dialog closed" -Type "Info"
}

# ====================================================={ Global Helper Functions for Multi-Row Tab System }=====================================================

# ---------------------- Select-TabGlobal Function ----------------------
### Never called. This is the start of tabbed window cod,e but I think it was never finished properly
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

    ## Hide all tab views
    if ($Script:AllTabs) {
      foreach ($t in $Script:AllTabs) {
        if ($t -and $t.View) { $t.View.Visible = $false }
      }
    }

    ## Show selected tab view
    if ($tab.View) {
      $tab.View.Visible = $true
      $Script:ActiveTab = $tab
      Debug-Log (": Tab view set to visible for: $tabText") -Type "Info"
    } else {
      Debug-Log (": WARNING - Tab has no View property") -Type "Warn"
      return
    }

    ## Update visual selection in both rows
    if ($Script:TabRows) {
      foreach ($row in $Script:TabRows) {
        if ($row) { $row.SelectedTab = $tab }
      }
      Debug-Log (": Updated row selections") -Type "Info"
    }

    ## Clear and add the selected tab's view to content host
    $contentHost.RemoveAll()
    $contentHost.Add($tab.View)
    $contentHost.SetNeedsDisplay()
    Debug-Log (": Added tab view to content host for: $tabText") -Type "Success"
  } catch {
    Debug-Log (": ERROR in Select-TabGlobal: $($_.Exception.Message)") -Type "Error"
    throw
  }
}

## ====================================================={ Main Dialog Function }=====================================================

function Show-UserPropertiesDialog {
  param($user)

  ## ---------------------- Safety Checks ----------------------
  if (-not $user) {
    Debug-Log ": User object is null" -Type "Warn"
    return
  }
  Debug-Log ": Show-UserPropertiesDialog starting for: $($user.Name)" -Type "Info"

  try {
    ## ---------------------- Buttons ----------------------
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply = [Terminal.Gui.Button]::new("Apply")

    ## Button handlers
    $btnOK.add_Clicked({
      Debug-Log ": OK clicked" -Type "Info"
      [Terminal.Gui.Application]::RequestStop()
    })

    $btnCancel.add_Clicked({
      Debug-Log ": Cancel clicked" -Type "Info"
      [Terminal.Gui.Application]::RequestStop()
    })

$btnApply.add_Clicked({
    Debug-Log ": Apply clicked - saving changes" -Type "Info"

    try {
        $changesMade = $false

        # Check SamAccountName change
        $newSamAccountName = $txtSamAccountName.Text.ToString().Trim()
        if ($newSamAccountName -ne $user.SamAccountName -and -not [string]::IsNullOrWhiteSpace($newSamAccountName)) {
            if ($Script:DemoMode) {
                $user.SamAccountName = $newSamAccountName
                Debug-Log ": Updated SamAccountName to '$newSamAccountName' (demo)" -Type "Success"
                $changesMade = $true
            } else {
                Set-ADUser -Identity $user.SamAccountName -SamAccountName $newSamAccountName -ErrorAction Stop
                $user.SamAccountName = $newSamAccountName
                Debug-Log ": Updated SamAccountName to '$newSamAccountName'" -Type "Success"
                $changesMade = $true
            }
        }

        # Check UserPrincipalName change
        $newUPN = $txtUserPrincipalName.Text.ToString().Trim()
        if ($newUPN -ne $user.UserPrincipalName -and -not [string]::IsNullOrWhiteSpace($newUPN)) {
            if ($Script:DemoMode) {
                $user.UserPrincipalName = $newUPN
                Debug-Log ": Updated UPN to '$newUPN' (demo)" -Type "Success"
                $changesMade = $true
            } else {
                Set-ADUser -Identity $user.SamAccountName -UserPrincipalName $newUPN -ErrorAction Stop
                $user.UserPrincipalName = $newUPN
                Debug-Log ": Updated UPN to '$newUPN'" -Type "Success"
                $changesMade = $true
            }
        }

        # Check other fields (DisplayName, Email, etc.)
        # ... your existing Apply logic ...

        if ($changesMade) {
            Show-Modal "Success" "Changes applied successfully"
        } else {
            Show-Modal "Info" "No changes to apply"
        }

    } catch {
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        Debug-Log ": Apply failed: $($_.Exception.Message)" -Type "Error"
    }
   })

    ## ---------------------- Dialog ----------------------
    $dlg = [Terminal.Gui.Dialog]::new("User Properties - $($user.Name)", 100, 40, $btnOK, $btnCancel, $btnApply)

    ## ---------------------- Standard TabView ----------------------
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(1)  # Leave room for buttons

    ## ==================== General Tab ====================
    Debug-Log ": Creating General tab" -Type "Info"
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"
    $generalView = [Terminal.Gui.View]::new()
    $generalView.X = 0; $generalView.Y = 0
    $generalView.Width = [Terminal.Gui.Dim]::Fill()
    $generalView.Height = [Terminal.Gui.Dim]::Fill()

    $y = 1
    ## Basic Info
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

    ## Phone numbers
    $lbl = [Terminal.Gui.Label]::new("Contact Information"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2
    $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
    $txtOfficePhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $txtOfficePhone.X=20; $txtOfficePhone.Y=$y; $txtOfficePhone.Width=30
    $generalView.Add($txtOfficePhone); $y+=1

    $lbl = [Terminal.Gui.Label]::new("Mobile:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
    $txtMobilePhone = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $txtMobilePhone.X=20; $txtMobilePhone.Y=$y; $txtMobilePhone.Width=30
    $generalView.Add($txtMobilePhone); $y+=1

    $generalTab.View = $generalView
    $tabView.AddTab($generalTab, $false)

    ## ==================== Account Tab ====================
    Debug-Log ": Creating Account tab" -Type "Info"
    $accountTab = [Terminal.Gui.TabView+Tab]::new()
    $accountTab.Text = "Account"
    $accountView = [Terminal.Gui.View]::new()
    $accountView.X = 0; $accountView.Y = 0
    $accountView.Width = [Terminal.Gui.Dim]::Fill()
    $accountView.Height = [Terminal.Gui.Dim]::Fill()

    $y = 1  # ← Make sure this is here!

    ## NEW: Logon Information at the TOP
    $lbl = [Terminal.Gui.Label]::new("Logon Information"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2

    $lbl = [Terminal.Gui.Label]::new("User logon name (UPN):"); $lbl.X=4; $lbl.Y=$y; $accountView.Add($lbl)
    $upn = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "" }
    $txtUserPrincipalName = [Terminal.Gui.TextField]::new($upn)
    $txtUserPrincipalName.X=35; $txtUserPrincipalName.Y=$y; $txtUserPrincipalName.Width=50
    $accountView.Add($txtUserPrincipalName); $y+=1

    $lbl = [Terminal.Gui.Label]::new("User logon name (Pre Win 2000):"); $lbl.X=4; $lbl.Y=$y; $accountView.Add($lbl)
    $txtSamAccountName = [Terminal.Gui.TextField]::new($user.SamAccountName ?? "")
    $txtSamAccountName.X=35; $txtSamAccountName.Y=$y; $txtSamAccountName.Width=50
    $accountView.Add($txtSamAccountName); $y+=2 # ← Extra spacing before next section

    ## Account Status
    $lbl = [Terminal.Gui.Label]::new("Account Status"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2
    $isEnabled = if ($user.PSObject.Properties['Enabled']) { $user.Enabled } else { -not $user.Disabled }
    $chkEnabled = [Terminal.Gui.CheckBox]::new("Account Enabled"); $chkEnabled.X=4; $chkEnabled.Y=$y; $chkEnabled.Checked=$isEnabled
    $accountView.Add($chkEnabled); $y+=1

    $isLocked = if ($user.PSObject.Properties['LockedOut']) { $user.LockedOut } else { $user.Locked ?? $false }
    $chkLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $chkLocked.X=4; $chkLocked.Y=$y; $chkLocked.Checked=$isLocked; $chkLocked.Enabled=$false
    $accountView.Add($chkLocked); $y+=2

    ## Password Settings
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

    ## Logon Information
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

    ## ==================== Address Tab ====================
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

    ## ==================== Profile Tab ====================
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

    ## Home Folder
    $lbl = [Terminal.Gui.Label]::new("Home Folder"); $lbl.X=2; $lbl.Y=$y; $profileView.Add($lbl); $y+=2

    $lbl = [Terminal.Gui.Label]::new("Home directory:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
    $txtHomeDirectory = [Terminal.Gui.TextField]::new($user.HomeDirectory ?? ""); $txtHomeDirectory.X=20; $txtHomeDirectory.Y=$y; $txtHomeDirectory.Width=70
    $profileView.Add($txtHomeDirectory); $y+=2

    $lbl = [Terminal.Gui.Label]::new("Home drive:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
    $txtHomeDrive = [Terminal.Gui.TextField]::new($user.HomeDrive ?? ""); $txtHomeDrive.X=20; $txtHomeDrive.Y=$y; $txtHomeDrive.Width=5
    $profileView.Add($txtHomeDrive); $y+=1

    $profileTab.View = $profileView
    $tabView.AddTab($profileTab, $false)

    ## ==================== Organization Tab ====================
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

  ## ==================== Member Of Tab ====================
  Debug-Log ": Creating Member Of tab" -Type "Info"
  $memberTab = [Terminal.Gui.TabView+Tab]::new()
  $memberTab.Text = "Member Of"
  $memberView = [Terminal.Gui.View]::new()
  $memberView.X = 0; $memberView.Y = 0
  $memberView.Width = [Terminal.Gui.Dim]::Fill()
  $memberView.Height = [Terminal.Gui.Dim]::Fill(3)
  $y = 1

  $lbl = [Terminal.Gui.Label]::new("Group Memberships:"); $lbl.X=2; $lbl.Y=$y; $memberView.Add($lbl); $y+=2

  ## Create ListView for groups
  $lstGroups = [Terminal.Gui.ListView]::new()
  $lstGroups.X = 2
  $lstGroups.Y = $y
  $lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
  $lstGroups.Height = 25  # Fixed height

  ## Get group list
  $groupList = @()
  if ($user.Groups) {
    $groupList = $user.Groups
  } elseif ($user.MemberOf) {
    ## Extract group names from DNs
    $groupList = $user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } }
  }

  if ($groupList.Count -gt 0) { $lstGroups.SetSource($groupList)
  } else { $lstGroups.SetSource(@("(No group memberships)"))
  }

  $memberView.Add($lstGroups)

  ## Add to Group button
  $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
  $btnAdd.X = 2
  $btnAdd.Y = 28  # y(1) + label(1) + list(25) + gap(1)
  $btnAdd.add_Clicked({
    Show-EditGroupMembershipDialog -User $user -OnUpdate {
    Debug-Log ": Refreshing group list after add" -Type "Info"

    ## Re-extract the group list from the updated user object
    $refreshedGroups = @()
    if ($user.Groups) {
      $refreshedGroups = $user.Groups
    } elseif ($user.MemberOf) {
      ## Extract group names from DNs
      $refreshedGroups = $user.MemberOf | ForEach-Object {
        if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
      }
    }

    ## Update the ListView
    if ($refreshedGroups.Count -gt 0) {
      $lstGroups.SetSource($refreshedGroups)
    } else {
      $lstGroups.SetSource(@("(No group memberships)"))
    }

    ## Also update the $groupList variable used by Remove button
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

  ## Get current group list
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

    ## Don't allow removal from "(No group memberships)" placeholder
    if ($selectedGroup -eq "(No group memberships)") {
      Show-Modal "Info" "No group selected"
      return
    }

    ## Confirm removal
    $confirmDlg = Show-Modal "Confirm Removal" "Remove $($user.Name) from group '$selectedGroup'?" -YesNo

    if ($confirmDlg -eq 0) {  ## Yes clicked
      try {
        if ($Script:DemoMode) {
          ## Demo mode - just update the list
          $user.Groups = $user.Groups | Where-Object { $_ -ne $selectedGroup }
          Debug-Log ": Removed $($user.Name) from group $selectedGroup (demo mode)" -Type "Success"
        } else {
          ## Production mode
          Remove-ADGroupMember -Identity $selectedGroup -Members $user.SamAccountName -Confirm:$false
          Debug-Log ": Removed $($user.Name) from group $selectedGroup" -Type "Success"
        }

        ## Refresh the list
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

## Helper function - dump all properties from the object
function Get-UserOutputLines($userObj) {
    $lines = @()
    $userObj.PSObject.Properties | ForEach-Object {
        $value = if ($_.Value -is [array]) {
            $_.Value -join ', '
        } elseif ($null -eq $_.Value) {
            ''
        } else {
            $_.Value.ToString()
        }
        $lines += "$($_.Name.PadRight(25)): $value"
    }
    return $lines
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

## Add TabView to dialog
$dlg.Add($tabView)

Debug-Log ": All tabs added, running dialog" -Type "Success"

## Run the dialog
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

# -------------------------{ Create New Object Wizard }-------------------------
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
  })

  $dlg.AddButton($btnCreate)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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
        Start-Sleep -Milliseconds 100

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

        Load-DomainData -domain $Script:CurrentDomain

        Debug-Log ("POST-LOAD: Users=$($Script:Users.Count), Objects=$($Script:ADObjects.Count), DCs=$($Script:DCs.Count)") -Type "Info"
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
          Load-DomainData -domain $Script:CurrentDomain
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

# -------------------------{ Change DC Dialog }-------------------------
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
        ## Demo mode - use rawDCs
        if ($Script:rawDCs) {
            $Script:rawDCs | Where-Object { $_.Domain -eq $Script:CurrentDomain }
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
    $lblCurrent = [Terminal.Gui.Label]::new("Current: $($Script:CurrentDC)")
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
        $currentMarker = if ($dc.Name -eq $Script:CurrentDC) { "► " } else { "  " }
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
        if ($availableDCs[$i].Name -eq $Script:CurrentDC) {
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
            $oldDC = $Script:CurrentDC
            $Script:CurrentDC = $selectedDC.Name

            Debug-Log ": Changed DC from $oldDC to $($Script:CurrentDC)" -Type "Success"

            ## Update status bar
            if ($Script:StatusItem) {
                $Script:StatusItem.Title = "Forest: $($Script:ForestName) | DC: $($Script:CurrentDC) | Domain: $($Script:CurrentDomain) | Objs: $($Script:ADObjects.Count)"
            }

            ## Refresh data from new DC
            Refresh-Data -domain $Script:CurrentDomain
            Build-Tree -domain $Script:CurrentDomain

            Show-Modal "DC Changed" "Successfully connected to $($Script:CurrentDC)"

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
    })
    $dlg.AddButton($btnCancel)

    ## Run dialog
    [Terminal.Gui.Application]::Run($dlg)
}

# -------------------------{ Tree Expand/Collapse }-------------------------
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
            if ($ChkDisabledOnly.Checked -and $objType -eq "User") { $filterStr = "Name -like '*$searchName*' -and Enabled -eq `$false" }
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
# ====================================================={ Example: How to Add Clipboard Buttons to Your UI }=====================================================

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
    Invoke-ADSearch -UserField $txtSearchName -DomainField $txtSearchDomain -ObjType $cmbSearchType -TabView $searchTabView -TxtOutput $txtSearchOutput -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab
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
    ## ---------------------- Buttons ----------------------
    $btnClose = [Terminal.Gui.Button]::new("Close")
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

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
    $btnCopyResults.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResults })
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
    $btnCopyQuery.add_Clicked({ Copy-LDAPQueryToClipboard -LdapFilter $txtLdapFilter })
    $advView.Add($btnCopyQuery)

    $btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
    $btnPasteQuery.add_Clicked({ Paste-LDAPQueryFromClipboard -LdapFilter $txtLdapFilter })
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
    $btnCopyResultsAdv.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResultsAdv })
    $advView.Add($btnCopyResultsAdv)
    $advTab.View = $advView
    $searchTabView.AddTab($advTab, $false)

    ## ==================== Wire up Search Buttons ====================

    ## Basic Search
    $btnSearch.add_Clicked({
      Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomain -ObjType $cmbObjectType -TabView $searchTabView -TxtOutput $txtResults -ChkDisabledOnly $chkDisabledOnly  -LdapFilter $txtLdapFilter -AdvTab $advTab
    })

    ## Advanced Search
    $btnSearchAdv.add_Clicked({
      Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomainAdv -ObjType $cmbObjectType -TabView $searchTabView -TxtOutput $txtResultsAdv -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab
    })

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

## -------------------------{ Refresh Tree Function }-------------------------
function Refresh-TreeData {
  Debug-Log (": Refreshing tree data...") -Type "Info"

  ## Show loading dialog
  $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."

  try {
    ## Reload domain data
    Load-DomainData -domain $Script:CurrentDomain

    ## Rebuild tree
    Build-Tree -domain $Script:CurrentDomain

    ## Add after Build-Tree calls:
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
    Debug-Log (": Tree refreshed successfully") -Type "Info"
  } finally {
    Close-LoadingDialog $loadingDlg
  }
  Show-Modal "Refreshed" "Active Directory data refreshed successfully"
}

function Show-DCPropertiesDialog {
  param($dc)

  ## ---------------------- Accept either DC object or DC name ----------------------
  if ($dc -is [string]) {
    $dcName = $dc
    Debug-Log ": Looking for DC: $dcName" -Type "Info"

    if (-not $Script:DCs) {
      Debug-Log ": Script:DCs is null or not initialized" -Type "Error"
      Show-Modal "Error" "Domain Controllers list is not loaded"
      return
    }

    Debug-Log ": Script:DCs has $($Script:DCs.Count) entries" -Type "Info"

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

  ## ==================== Tab Definitions ====================

  # General Tab
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      ## Basic Information
      $lbl = [Terminal.Gui.Label]::new("Domain Controller Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $lbl = [Terminal.Gui.Label]::new($dc.Name ?? "Unknown"); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1

      $hostname = if ($dc.HostName) { $dc.HostName } elseif ($dc.DNSHostName) { $dc.DNSHostName } else { $dc.Name }
      $lbl = [Terminal.Gui.Label]::new("Hostname:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $lbl = [Terminal.Gui.Label]::new($hostname); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1

      if ($dc.Site) {
        $lbl = [Terminal.Gui.Label]::new("Site:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Site); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.Domain) {
        $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Domain); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.Forest) {
        $lbl = [Terminal.Gui.Label]::new("Forest:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Forest); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.Location) {
        $lbl = [Terminal.Gui.Label]::new("Location:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Location); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      $y+=1

      ## Network Information
      $lbl = [Terminal.Gui.Label]::new("Network"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($dc.IPv4Address -or $dc.IPAddress) {
        $ip = if ($dc.IPv4Address) { $dc.IPv4Address } else { $dc.IPAddress }
        $lbl = [Terminal.Gui.Label]::new("IPv4 Address:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($ip); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.IPv6Address) {
        $lbl = [Terminal.Gui.Label]::new("IPv6 Address:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.IPv6Address); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      $y+=1

      ## Operating System
      $lbl = [Terminal.Gui.Label]::new("Operating System"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($dc.OperatingSystem -or $dc.OS) {
        $os = if ($dc.OperatingSystem) { $dc.OperatingSystem } else { $dc.OS }
        $lbl = [Terminal.Gui.Label]::new("OS:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($os); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.OperatingSystemVersion) {
        $lbl = [Terminal.Gui.Label]::new("OS Version:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.OperatingSystemVersion); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      $y+=1

      ## Capabilities
      $lbl = [Terminal.Gui.Label]::new("Capabilities"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $isGC = if ($null -ne $dc.IsGlobalCatalog) { $dc.IsGlobalCatalog.ToString() } else { "Unknown" }
      $lbl = [Terminal.Gui.Label]::new("Global Catalog:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $lbl = [Terminal.Gui.Label]::new($isGC); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1

      if ($null -ne $dc.IsReadOnly) {
        $lbl = [Terminal.Gui.Label]::new("Read-Only:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.IsReadOnly.ToString()); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($null -ne $dc.Enabled) {
        $lbl = [Terminal.Gui.Label]::new("Enabled:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Enabled.ToString()); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  # Roles Tab
  $rolesTab = @{
    Name = "Roles"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("FSMO Roles"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      ## Get FSMO roles from either FSMORoles or OperationMasterRoles
      $fsmoRoles = @()
      if ($dc.PSObject.Properties['FSMORoles'] -and $dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) {
        $fsmoRoles = $dc.FSMORoles
      } elseif ($dc.PSObject.Properties['OperationMasterRoles'] -and $dc.OperationMasterRoles -and $dc.OperationMasterRoles.Count -gt 0) {
        $fsmoRoles = $dc.OperationMasterRoles
      }

      if ($fsmoRoles.Count -gt 0) {
        foreach ($role in $fsmoRoles) {
          $lbl = [Terminal.Gui.Label]::new("• $role"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      } else {
        $lbl = [Terminal.Gui.Label]::new("(None)"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  # Replication Tab
  $replicationTab = @{
    Name = "Replication"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Replication Status"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($dc.ReplicationHealth) {
        $lbl = [Terminal.Gui.Label]::new("Health:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.ReplicationHealth); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($dc.LastReplication) {
        $lastRep = $dc.LastReplication.ToString('yyyy-MM-dd HH:mm:ss')
        $lbl = [Terminal.Gui.Label]::new("Last Replication:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($lastRep); $lbl.X=25; $lbl.Y=$y; $view.Add($lbl); $y+=2
      }

      ## Replication Partners
      if ($dc.ReplicationPartners -and $dc.ReplicationPartners.Count -gt 0) {
        $lbl = [Terminal.Gui.Label]::new("Replication Partners:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        foreach ($partner in $dc.ReplicationPartners) {
          $lbl = [Terminal.Gui.Label]::new("• $partner"); $lbl.X=6; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      }
    }
  }

  # Services Tab
  $servicesTab = @{
    Name = "Services"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      ## Services Status
      if ($dc.Services) {
        $lbl = [Terminal.Gui.Label]::new("Service Status"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
        foreach ($service in $dc.Services.Keys | Sort-Object) {
          $status = $dc.Services[$service]
          $lbl = [Terminal.Gui.Label]::new("${service}:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($status); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
        $y+=1
      }

      ## Boot Time
      if ($dc.LastBoot -or $dc.LastBootUpTime) {
        $lbl = [Terminal.Gui.Label]::new("System Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
        $lastBoot = if ($dc.LastBootUpTime) { $dc.LastBootUpTime } else { $dc.LastBoot }
        if ($lastBoot) {
          $bootTime = $lastBoot.ToString('yyyy-MM-dd HH:mm:ss')
          $uptime = (Get-Date) - $lastBoot
          $uptimeStr = "$($uptime.Days) days, $($uptime.Hours) hours"
          $lbl = [Terminal.Gui.Label]::new("Last Boot:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($bootTime); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          $lbl = [Terminal.Gui.Label]::new("Uptime:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($uptimeStr); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      }
    }
  }

  # Disk Space Tab
  $diskTab = @{
    Name = "Disk Space"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      if ($dc.DiskSpace) {
        $lbl = [Terminal.Gui.Label]::new("Disk Usage"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
        foreach ($drive in $dc.DiskSpace.Keys | Sort-Object) {
          $diskInfo = $dc.DiskSpace[$drive]
          $lbl = [Terminal.Gui.Label]::new("Drive ${drive}"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

          if ($diskInfo.Total) {
            $lbl = [Terminal.Gui.Label]::new("  Total:"); $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Total); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }

          if ($diskInfo.Used) {
            $lbl = [Terminal.Gui.Label]::new("  Used:"); $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Used); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }

          if ($diskInfo.Free) {
            $lbl = [Terminal.Gui.Label]::new("  Free:"); $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Free); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }

          if ($null -ne $diskInfo.PercentFree) {
            $lbl = [Terminal.Gui.Label]::new("  % Free:"); $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new("$($diskInfo.PercentFree)%"); $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }
          $y+=1
        }
      } else {
        $lbl = [Terminal.Gui.Label]::new("(No disk space information available)"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

# ===================== Search / Lookup (Domain Controller) Tab =====================
$searchDCPropertiesTab = @{
  Name = "Search/Lookup"
  Builder = {
    param($view, $dc, $state)

    $y = 1

    # ---- Domain ----
    $lblSearchDomain = [Terminal.Gui.Label]::new("Domain:")
    $lblSearchDomain.X = 2
    $lblSearchDomain.Y = $y
    $view.Add($lblSearchDomain)

    $state.txtSearchDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain ?? "")
    $state.txtSearchDomain.X = 15
    $state.txtSearchDomain.Y = $y
    $state.txtSearchDomain.Width = 30
    $view.Add($state.txtSearchDomain)
    $y += 2

    # ---- DC Name ----
    $lblSearchName = [Terminal.Gui.Label]::new("DC:")
    $lblSearchName.X = 2
    $lblSearchName.Y = $y
    $view.Add($lblSearchName)

    $state.txtSearchUser = [Terminal.Gui.TextField]::new($dc.Name ?? "")
    $state.txtSearchUser.X = 15
    $state.txtSearchUser.Y = $y
    $state.txtSearchUser.Width = 30
    $view.Add($state.txtSearchUser)
    $y += 2

    # ---- Search Type ----
    $lblSearchType = [Terminal.Gui.Label]::new("Type:")
    $lblSearchType.X = 2
    $lblSearchType.Y = $y
    $view.Add($lblSearchType)

    $state.cmbSearchType = [Terminal.Gui.ComboBox]::new()
    $state.cmbSearchType.X = 15
    $state.cmbSearchType.Y = $y
    $state.cmbSearchType.Width = 20
    $state.cmbSearchType.SetSource(@("Domain Controller","Computer","OU"))
    $state.cmbSearchType.SelectedItem = 0
    $view.Add($state.cmbSearchType)
    $y += 2

    # ---- Filter Results ----
    $lblSearchFilter = [Terminal.Gui.Label]::new("Filter Results:")
    $lblSearchFilter.X = 48
    $lblSearchFilter.Y = 1
    $view.Add($lblSearchFilter)

    $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
    $state.txtSearchFilter.X = 62
    $state.txtSearchFilter.Y = 1
    $state.txtSearchFilter.Width = 20
    $view.Add($state.txtSearchFilter)

    $state.txtSearchFilter.add_TextChanged({
      if ($state.currentSearchOutputLines) {
        $search = $state.txtSearchFilter.Text.ToString().Trim()
        if ($search) {
          $state.txtSearchOutput.Text =
            ($state.currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
        }
        else {
          $state.txtSearchOutput.Text = $state.currentSearchOutputLines -join "`n"
        }
      }
    }.GetNewClosure())

    # ---- Results ----
    $lblSearchResult = [Terminal.Gui.Label]::new("Results:")
    $lblSearchResult.X = 2
    $lblSearchResult.Y = $y
    $view.Add($lblSearchResult)
    $y += 1

    $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
    $state.txtSearchOutput.X = 2
    $state.txtSearchOutput.Y = $y
    $state.txtSearchOutput.Width  = [Terminal.Gui.Dim]::Fill(2)
    $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(4)
    $state.txtSearchOutput.ReadOnly = $true
    $state.txtSearchOutput.WordWrap = $false
    $view.Add($state.txtSearchOutput)

    # ---- DC Status ----
    $state.chkSearchLocked = [Terminal.Gui.CheckBox]::new("Global Catalog")
    $state.chkSearchLocked.X = 2
    $state.chkSearchLocked.Y = [Terminal.Gui.Pos]::Bottom($state.txtSearchOutput) + 1
    $state.chkSearchLocked.CanFocus = $true
    $view.Add($state.chkSearchLocked)

    $state.chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Read-Only DC")
    $state.chkSearchDisabled.X = 2
    $state.chkSearchDisabled.Y = [Terminal.Gui.Pos]::Bottom($state.chkSearchLocked) + 1
    $state.chkSearchDisabled.CanFocus = $true
    $view.Add($state.chkSearchDisabled)

    # ---- Search Button (UI only) ----
    $btnDoSearch = [Terminal.Gui.Button]::new("Search")
    $btnDoSearch.X = 48
    $btnDoSearch.Y = 3
    $view.Add($btnDoSearch)

    # ---- Auto-populate DC properties ----
    [Terminal.Gui.Application]::MainLoop.Invoke({
      if ($dc) {

        $lines = @()
        $dc.PSObject.Properties | ForEach-Object {
          $value = if ($_.Value -is [array]) {
            $_.Value -join ', '
          }
          elseif ($null -eq $_.Value) {
            ''
          }
          else {
            $_.Value.ToString()
          }

          $lines += "$($_.Name.PadRight(25)): $value"
        }

        $state.txtSearchOutput.Text = $lines -join "`n"
        $state.currentSearchOutputLines = $lines

        # Best-effort flags (safe for demo objects)
        $state.chkSearchLocked.Checked   = [bool]($dc.IsGlobalCatalog)
        $state.chkSearchDisabled.Checked = [bool]($dc.IsReadOnly)
      }
    }.GetNewClosure())
  }
}

  ## ==================== Create and Show Dialog ====================
  # DC properties are read-only, no Apply logic needed
  $tabs = @($generalTab, $rolesTab, $replicationTab, $servicesTab, $diskTab, $searchDCPropertiesTab)

  Debug-Log ": All DC tabs added, running dialog" -Type "Success"

  New-PropertiesDialog -Title "Domain Controller Properties - $($dc.Name)" -Width 110 -Height 35 -Tabs $tabs -Data $dc
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
              if ($Script:rawUsers) {
                $Script:rawUsers = $Script:rawUsers | Where-Object { $_.Name -ne $name }
              }
            }
            'Group' {
              if ($Script:Groups) {
                $Script:Groups = $Script:Groups | Where-Object { $_.Name -ne $name }
              }
              if ($Script:rawDemoGroups) {
                $Script:rawDemoGroups = $Script:rawDemoGroups | Where-Object { $_.Name -ne $name }
              }
            }
            'Computer' {
              if ($Script:Computers) {
                $Script:Computers = $Script:Computers | Where-Object { $_.Name -ne $name }
              }
            }
            'OU' {
              if ($Script:rawOUs) {
                $Script:rawOUs = $Script:rawOUs | Where-Object { $_.Name -ne $name }
              }
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
        if ($Script:FilterStatusLabel) {
          Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
        }
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
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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
        $isUser = $obj.PSObject.Properties.Match('SamAccountName').Count -gt 0 -and
                  -not $obj.PSObject.Properties.Match('ComputerType')
        $isComputer = $obj.PSObject.Properties.Match('ComputerType').Count -gt 0

        if ($isUser -or $isComputer) {
            $validObjects += $obj
        }
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
                    if ($objectType -eq 'User') {
                        Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
                    }
                } else {
                    if ($objectType -eq 'User') {
                        Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
                    } else {
                        Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
                    }
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
    if ($Script:FilterStatusLabel) {
        Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
    }

    Debug-Log ": Bulk $Action completed - $successCount succeeded, $failCount failed" -Type "Success"
}

function New-PropertiesDialog {
    <#
    .SYNOPSIS
    Creates a standard properties dialog with tabs

    .PARAMETER Title
    Dialog title (e.g., "User Properties - Jim Kerr")

    .PARAMETER Width
    Dialog width (default 100)

    .PARAMETER Height
    Dialog height (default 40)

    .PARAMETER Tabs
    Array of tab definitions. Each tab is a hashtable with:
    - Name: Tab title
    - Builder: Scriptblock that receives ($view, $data, $sharedState)

    .PARAMETER Data
    The object being edited (user, group, computer, etc.)

    .PARAMETER OnApply
    Scriptblock to execute when Apply is clicked

    .PARAMETER OnOK
    Scriptblock to execute when OK is clicked (default: just close)

    .EXAMPLE
    $tabs = @(
        @{
            Name = "General"
            Builder = {
                param($view, $user, $state)
                $y = 1
                $lbl = [Terminal.Gui.Label]::new("Name:")
                $lbl.X = 2; $lbl.Y = $y
                $view.Add($lbl)

                $state.txtName = [Terminal.Gui.TextField]::new($user.Name)
                $state.txtName.X = 20; $state.txtName.Y = $y
                $view.Add($state.txtName)
            }
        }
    )

    New-PropertiesDialog -Title "User Properties" -Tabs $tabs -Data $user -OnApply {
        param($data, $state)
        # Save changes using $state.txtName.Text, etc.
    }
    #>
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
        [scriptblock]$OnOK
    )

    try {
        # Shared state for controls (accessible across tabs and buttons)
        $sharedState = @{}

        # Create buttons
        $btnOK = [Terminal.Gui.Button]::new("OK")
        $btnCancel = [Terminal.Gui.Button]::new("Cancel")
        $btnApply = [Terminal.Gui.Button]::new("Apply")

        # Button handlers
        $btnOK.add_Clicked({
            Debug-Log ": OK clicked" -Type "Info"
            if ($OnOK) {
                & $OnOK $Data $sharedState
            }
            [Terminal.Gui.Application]::RequestStop()
        }.GetNewClosure())

        $btnCancel.add_Clicked({
            Debug-Log ": Cancel clicked" -Type "Info"
            [Terminal.Gui.Application]::RequestStop()
        })

        $btnApply.add_Clicked({
            Debug-Log ": Apply clicked" -Type "Info"
            if ($OnApply) {
                try {
                    & $OnApply $Data $sharedState
                } catch {
                    Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
                    Debug-Log ": Apply failed: $($_.Exception.Message)" -Type "Error"
                }
            }
        }.GetNewClosure())

        # Create dialog
        $dlg = [Terminal.Gui.Dialog]::new($Title, $Width, $Height, $btnOK, $btnCancel, $btnApply)

        # Create TabView
        $tabView = [Terminal.Gui.TabView]::new()
        $tabView.X = 0
        $tabView.Y = 0
        $tabView.Width = [Terminal.Gui.Dim]::Fill()
        $tabView.Height = [Terminal.Gui.Dim]::Fill(1)  # Leave room for dialog buttons

        # Build each tab
        foreach ($tabDef in $Tabs) {
            Debug-Log ": Creating tab: $($tabDef.Name)" -Type "Info"

            # Create tab
            $tab = [Terminal.Gui.TabView+Tab]::new()
            $tab.Text = $tabDef.Name

            # Create view
            $view = [Terminal.Gui.View]::new()
            $view.X = 0; $view.Y = 0
            $view.Width = [Terminal.Gui.Dim]::Fill()
            $view.Height = [Terminal.Gui.Dim]::Fill()

            # Call builder to populate view
            if ($tabDef.Builder) {
                try {
                    & $tabDef.Builder $view $Data $sharedState
                } catch {
                    Debug-Log ": Error building tab $($tabDef.Name): $($_.Exception.Message)" -Type "Error"
                }
            }

            # Attach and add
            $tab.View = $view
            $tabView.AddTab($tab, $false)
        }

        # Add TabView to dialog
        $dlg.Add($tabView)

        Debug-Log ": All tabs added, running dialog" -Type "Success"

        # Run dialog
        [Terminal.Gui.Application]::Run($dlg)

        Debug-Log ": Dialog closed normally" -Type "Info"

    } catch {
        Debug-Log ": Exception in New-PropertiesDialog: $($_.Exception.Message)" -Type "Error"
        Debug-Log ": Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        Show-Modal "Error" "Failed to display dialog:`n$($_.Exception.Message)"
    }
}

function Show-GroupPropertiesDialog {
  param($group)

  ## ---------------------- Safety Checks ----------------------
  if (-not $group) {
    Debug-Log ": Group object is null" -Type "Warn"
    return
  }
  Debug-Log ": Show-GroupPropertiesDialog starting for: $($group.Name)" -Type "Info"

  ## ==================== Tab Definitions ====================

  # General Tab
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $group, $state)
      $y = 1
      $lbl = [Terminal.Gui.Label]::new("Group Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Group name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtName = [Terminal.Gui.Label]::new($group.Name ?? ""); $state.txtName.X=20; $state.txtName.Y=$y
      $view.Add($state.txtName); $y+=1
      $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDescription = [Terminal.Gui.TextField]::new($group.Description ?? ""); $state.txtDescription.X=20; $state.txtDescription.Y=$y; $state.txtDescription.Width=60
      $view.Add($state.txtDescription); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Group Details"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Type:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.lblType = [Terminal.Gui.Label]::new($group.Type ?? "Security"); $state.lblType.X=20; $state.lblType.Y=$y
      $view.Add($state.lblType); $y+=1
      $lbl = [Terminal.Gui.Label]::new("Scope:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.lblScope = [Terminal.Gui.Label]::new($group.Scope ?? "Global"); $state.lblScope.X=20; $state.lblScope.Y=$y
      $view.Add($state.lblScope); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Contact Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Email:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtEmail = [Terminal.Gui.TextField]::new($group.Email ?? ""); $state.txtEmail.X=20; $state.txtEmail.Y=$y; $state.txtEmail.Width=60
      $view.Add($state.txtEmail); $y+=1
      $lbl = [Terminal.Gui.Label]::new("Managed by:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtManagedBy = [Terminal.Gui.TextField]::new($group.ManagedBy ?? ""); $state.txtManagedBy.X=20; $state.txtManagedBy.Y=$y; $state.txtManagedBy.Width=60
      $view.Add($state.txtManagedBy); $y+=1
    }
  }
    $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
    $btnAuditLog.X = 2
    $btnAuditLog.Y = $y  # Position appropriately
    $btnAuditLog.add_Clicked({
      Show-AuditLogDialog -Object $user -ObjectType 'Group'  # Adjust for group/computer
    }.GetNewClosure())
    $view.Add($btnAuditLog)

  # Members Tab - (existing code truncated for brevity - no changes)
  $membersTab = @{
    Name = "Members"
    Builder = {
      param($view, $group, $state)
      $y = 1
      $lbl = [Terminal.Gui.Label]::new("Group Members:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.lstMembers = [Terminal.Gui.ListView]::new()
      $state.lstMembers.X = 2; $state.lstMembers.Y = $y; $state.lstMembers.Width = [Terminal.Gui.Dim]::Fill(2); $state.lstMembers.Height = 25
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
      if ($state.memberList.Count -gt 0) { $state.lstMembers.SetSource($state.memberList) }
      else { $state.lstMembers.SetSource(@("(No members)")) }
      $view.Add($state.lstMembers)
      ## ... (Add/Remove buttons - same as original)
    }
  }

  # Members Report Tab - NEW!
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

      ## Header & Statistics
      $lblHeader = [Terminal.Gui.Label]::new("Membership Report: $($group.Name)")
      $lblHeader.X = 2; $lblHeader.Y = $y; $view.Add($lblHeader); $y += 2

      $totalMembers = $memberDetails.Count
      $enabledMembers = ($memberDetails | Where-Object { $_.Enabled -eq $true }).Count
      $disabledMembers = ($memberDetails | Where-Object { $_.Enabled -eq $false }).Count

      $lblStats = [Terminal.Gui.Label]::new("Total Members: $totalMembers (Enabled: $enabledMembers, Disabled: $disabledMembers)")
      $lblStats.X = 2; $lblStats.Y = $y; $view.Add($lblStats); $y += 2

      ## Member list with details
      $lstReport = [Terminal.Gui.ListView]::new()
      $lstReport.X = 2; $lstReport.Y = $y; $lstReport.Width = [Terminal.Gui.Dim]::Fill(2); $lstReport.Height = 22

      $displayItems = @()
      foreach ($member in ($memberDetails | Sort-Object Name)) {
        $statusIcon = if ($member.Enabled) { "✓" } else { "⊗" }
        $displayItems += "[$statusIcon] $($member.Name) | $($member.Department ?? 'N/A') | $($member.Title ?? 'N/A') | $($member.Email ?? 'N/A')"
      }

      if ($displayItems.Count -eq 0) { $displayItems = @("(No members)") }
      $lstReport.SetSource($displayItems); $view.Add($lstReport)

      $y = 26

      ## Export Full Report button
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

      ## Compare with Another Group button
      $btnCompare = [Terminal.Gui.Button]::new("Compare with Group...")
      $btnCompare.X = 25; $btnCompare.Y = $y
      $btnCompare.add_Clicked({
        ## Show group selector
        $compareDlg = [Terminal.Gui.Dialog]::new("Compare Groups", 60, 20)
        $lblInfo = [Terminal.Gui.Label]::new("Select group to compare with $($group.Name):")
        $lblInfo.X = 2; $lblInfo.Y = 1; $compareDlg.Add($lblInfo)

        $otherGroups = if ($Script:DemoMode) {
          $Script:Groups | Where-Object { $_.Name -ne $group.Name } | Select-Object -ExpandProperty Name | Sort-Object
        } else {
          try { Get-ADGroup -Filter * | Where-Object { $_.Name -ne $group.Name } | Select-Object -ExpandProperty Name | Sort-Object }
          catch { @() }
        }

        $lstGroups = [Terminal.Gui.ListView]::new($otherGroups)
        $lstGroups.X = 2; $lstGroups.Y = 3; $lstGroups.Width = 54; $lstGroups.Height = 12
        $compareDlg.Add($lstGroups)

        $btnSelect = [Terminal.Gui.Button]::new("Compare")
        $btnSelect.add_Clicked({
          if ($lstGroups.SelectedItem -ge 0) {
            $compareGroupName = $otherGroups[$lstGroups.SelectedItem]
            [Terminal.Gui.Application]::RequestStop()

            ## Perform comparison
            $group1Members = $memberDetails | Select-Object -ExpandProperty SamAccountName
            $group2Members = if ($Script:DemoMode) {
              $Script:Users | Where-Object { $_.Groups -contains $compareGroupName } | Select-Object -ExpandProperty SamAccountName
            } else {
              try { Get-ADGroupMember -Identity $compareGroupName -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName }
              catch { @() }
            }

            ## Calculate differences
            $inBoth = $group1Members | Where-Object { $group2Members -contains $_ }
            $onlyInGroup1 = $group1Members | Where-Object { $group2Members -notcontains $_ }
            $onlyInGroup2 = $group2Members | Where-Object { $group1Members -notcontains $_ }

            ## Show results
            $resultMsg = "Comparison: $($group.Name) vs $compareGroupName`n`n"
            $resultMsg += "In both groups: $($inBoth.Count)`n"
            $resultMsg += "Only in $($group.Name): $($onlyInGroup1.Count)`n"
            $resultMsg += "Only in ${compareGroupName}: $($onlyInGroup2.Count)"
            Show-Modal "Group Comparison" $resultMsg

            Debug-Log ": Compared groups - Both: $($inBoth.Count), Only1: $($onlyInGroup1.Count), Only2: $($onlyInGroup2.Count)" -Type "Info"
          }
        }.GetNewClosure())
        $compareDlg.AddButton($btnSelect)

        $btnCancel = [Terminal.Gui.Button]::new("Cancel")
        $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
        $compareDlg.AddButton($btnCancel)

        [Terminal.Gui.Application]::Run($compareDlg)
      }.GetNewClosure())
      $view.Add($btnCompare)
    }
  }

  ## ==================== Apply Logic (unchanged) ====================
  $applyLogic = {
    param($group, $state)
    Debug-Log ": Apply clicked - saving changes" -Type "Info"
    try {
      $changesMade = $false
      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $group.Description) {
          if ($Script:DemoMode) { $group.Description = $newDescription; $changesMade = $true }
          else { Set-ADGroup -Identity $group.Name -Description $newDescription -ErrorAction Stop; $group.Description = $newDescription; $changesMade = $true }
        }
      }
      if ($state.txtEmail) {
        $newEmail = $state.txtEmail.Text.ToString().Trim()
        if ($newEmail -ne $group.Email) {
          if ($Script:DemoMode) { $group.Email = $newEmail; $changesMade = $true }
          else { Set-ADGroup -Identity $group.Name -Replace @{mail=$newEmail} -ErrorAction Stop; $group.Email = $newEmail; $changesMade = $true }
        }
      }
      if ($state.txtManagedBy) {
        $newManagedBy = $state.txtManagedBy.Text.ToString().Trim()
        if ($newManagedBy -ne $group.ManagedBy) {
          if ($Script:DemoMode) { $group.ManagedBy = $newManagedBy; $changesMade = $true }
          else { Set-ADGroup -Identity $group.Name -ManagedBy $newManagedBy -ErrorAction Stop; $group.ManagedBy = $newManagedBy; $changesMade = $true }
        }
      }
      if ($changesMade) { Show-Modal "Success" "Changes applied successfully" }
      else { Show-Modal "Info" "No changes to apply" }
    } catch {
      Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
      Debug-Log ": Apply failed: $($_.Exception.Message)" -Type "Error"
    }
  }

  ## ==================== Create and Show Dialog ====================
  $tabs = @($generalTab, $membersTab, $reportTab)
  New-PropertiesDialog -Title "Group Properties - $($group.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $group -OnApply $applyLogic
}

# ----------------------- Show Computer Properties ----------------------
function Show-ComputerPropertiesDialog {
  param([string]$computerName)

  ## ---------------------- Safety Checks ----------------------
  Debug-Log ": Showing computer properties for: $computerName" -Type "Info"
  $computer = $Script:Computers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1

  if (-not $computer) {
    Debug-Log ": Computer NOT found in Script:Computers for name: $computerName" -Type "Info"
    Show-Modal "Not Found" "Computer '$computerName' not found"
    return
  }

  Debug-Log ": Computer found: $($computer.Name)" -Type "Info"

  ## ==================== Tab Definitions ====================

  # General Tab
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Computer Information"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtName = [Terminal.Gui.TextField]::new($computer.Name ?? ""); $state.txtName.X=25; $state.txtName.Y=$y
      $view.Add($state.txtName); $y+=1

      $lbl = [Terminal.Gui.Label]::new("DNS Host Name:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDNS = [Terminal.Gui.Label]::new($computer.DNSHostName ?? ""); $state.txtDNS.X=25; $state.txtDNS.Y=$y
      $view.Add($state.txtDNS); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDomain = [Terminal.Gui.Label]::new($computer.Domain ?? ""); $state.txtDomain.X=25; $state.txtDomain.Y=$y
      $view.Add($state.txtDomain); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDescription = [Terminal.Gui.TextField]::new($computer.Description ?? ""); $state.txtDescription.X=25; $state.txtDescription.Y=$y; $state.txtDescription.Width=65
      $view.Add($state.txtDescription); $y+=2

      ## Operating System
      $lbl = [Terminal.Gui.Label]::new("Operating System"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("OS:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtOS = [Terminal.Gui.Label]::new($computer.OperatingSystem ?? ""); $state.txtOS.X=25; $state.txtOS.Y=$y
      $view.Add($state.txtOS); $y+=1

      $lbl = [Terminal.Gui.Label]::new("Version:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtOSVer = [Terminal.Gui.Label]::new($computer.OperatingSystemVersion ?? ""); $state.txtOSVer.X=25; $state.txtOSVer.Y=$y
      $view.Add($state.txtOSVer); $y+=1

      if ($computer.OperatingSystemServicePack) {
        $lbl = [Terminal.Gui.Label]::new("Service Pack:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $state.txtSP = [Terminal.Gui.Label]::new($computer.OperatingSystemServicePack); $state.txtSP.X=25; $state.txtSP.Y=$y
        $view.Add($state.txtSP); $y+=1
      }

      ## Network
      $y+=1
      $lbl = [Terminal.Gui.Label]::new("Network"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("IPv4 Address:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtIPv4 = [Terminal.Gui.Label]::new($computer.IPv4Address ?? ""); $state.txtIPv4.X=25; $state.txtIPv4.Y=$y
      $view.Add($state.txtIPv4); $y+=1

      if ($computer.IPv6Address) {
        $lbl = [Terminal.Gui.Label]::new("IPv6 Address:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $state.txtIPv6 = [Terminal.Gui.Label]::new($computer.IPv6Address); $state.txtIPv6.X=25; $state.txtIPv6.Y=$y
        $view.Add($state.txtIPv6); $y+=1
      }

      $lbl = [Terminal.Gui.Label]::new("Location:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtLocation = [Terminal.Gui.TextField]::new($computer.Location ?? ""); $state.txtLocation.X=25; $state.txtLocation.Y=$y; $state.txtLocation.Width=65
      $view.Add($state.txtLocation); $y+=1
    }
  }

  # Account Tab
  $accountTab = @{
    Name = "Account"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      ## Account Status
      $lbl = [Terminal.Gui.Label]::new("Account Status"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.chkEnabled = [Terminal.Gui.CheckBox]::new("Computer Account Enabled"); $state.chkEnabled.X=4; $state.chkEnabled.Y=$y; $state.chkEnabled.Checked=($computer.Enabled??$true)
      $view.Add($state.chkEnabled); $y+=1

      ## Password Settings
      $lbl = [Terminal.Gui.Label]::new("Password Settings"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired"); $state.chkPasswordExpired.X=4; $state.chkPasswordExpired.Y=$y; $state.chkPasswordExpired.Checked=($computer.PasswordExpired??$false); $state.chkPasswordExpired.Enabled=$false
      $view.Add($state.chkPasswordExpired); $y+=1

      $state.chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires"); $state.chkPasswordNeverExpires.X=4; $state.chkPasswordNeverExpires.Y=$y; $state.chkPasswordNeverExpires.Checked=($computer.PasswordNeverExpires??$false); $state.chkPasswordNeverExpires.Enabled=$false
      $view.Add($state.chkPasswordNeverExpires); $y+=1

      $state.chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("Cannot change password"); $state.chkCannotChangePassword.X=4; $state.chkCannotChangePassword.Y=$y; $state.chkCannotChangePassword.Checked=($computer.CannotChangePassword??$false); $state.chkCannotChangePassword.Enabled=$false
      $view.Add($state.chkCannotChangePassword); $y+=1

      $state.chkPasswordNotRequired = [Terminal.Gui.CheckBox]::new("Password not required"); $state.chkPasswordNotRequired.X=4; $state.chkPasswordNotRequired.Y=$y; $state.chkPasswordNotRequired.Checked=($computer.PasswordNotRequired??$false); $state.chkPasswordNotRequired.Enabled=$false
      $view.Add($state.chkPasswordNotRequired); $y+=2

      ## Account Details
      $lbl = [Terminal.Gui.Label]::new("Account Details"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $lbl = [Terminal.Gui.Label]::new("SAM Account:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      $state.txtSAM = [Terminal.Gui.Label]::new($computer.SamAccountName ?? ""); $state.txtSAM.X=25; $state.txtSAM.Y=$y
      $view.Add($state.txtSAM); $y+=1

      if ($computer.PasswordLastSet) {
        $lbl = [Terminal.Gui.Label]::new("Password last set: " + $computer.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.LastLogonDate) {
        $lbl = [Terminal.Gui.Label]::new("Last logon: " + $computer.LastLogonDate.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.logonCount) {
        $lbl = [Terminal.Gui.Label]::new("Logon count: " + $computer.logonCount); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.AccountExpirationDate) {
        $lbl = [Terminal.Gui.Label]::new("Account expires: " + $computer.AccountExpirationDate.ToString('yyyy-MM-dd')); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }
    }
  }

    $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
    $btnAuditLog.X = 2
    $btnAuditLog.Y = $y  # Position appropriately
    $btnAuditLog.add_Clicked({
      Show-AuditLogDialog -Object $user -ObjectType 'Computer'  # Adjust for group/computer
    }.GetNewClosure())
    $view.Add($btnAuditLog)

  # Security Tab
  $securityTab = @{
    Name = "Security"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Delegation Settings"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.chkTrustedForDelegation = [Terminal.Gui.CheckBox]::new("Trusted for delegation"); $state.chkTrustedForDelegation.X=4; $state.chkTrustedForDelegation.Y=$y; $state.chkTrustedForDelegation.Checked=($computer.TrustedForDelegation??$false); $state.chkTrustedForDelegation.Enabled=$false
      $view.Add($state.chkTrustedForDelegation); $y+=1

      $state.chkTrustedToAuth = [Terminal.Gui.CheckBox]::new("Trusted to authenticate for delegation"); $state.chkTrustedToAuth.X=4; $state.chkTrustedToAuth.Y=$y; $state.chkTrustedToAuth.Checked=($computer.TrustedToAuthForDelegation??$false); $state.chkTrustedToAuth.Enabled=$false
      $view.Add($state.chkTrustedToAuth); $y+=1

      $state.chkAccountNotDelegated = [Terminal.Gui.CheckBox]::new("Account not delegated"); $state.chkAccountNotDelegated.X=4; $state.chkAccountNotDelegated.Y=$y; $state.chkAccountNotDelegated.Checked=($computer.AccountNotDelegated??$false); $state.chkAccountNotDelegated.Enabled=$false
      $view.Add($state.chkAccountNotDelegated); $y+=1

      $state.chkNoPreAuth = [Terminal.Gui.CheckBox]::new("Does not require Kerberos preauthentication"); $state.chkNoPreAuth.X=4; $state.chkNoPreAuth.Y=$y; $state.chkNoPreAuth.Checked=($computer.DoesNotRequirePreAuth??$false); $state.chkNoPreAuth.Enabled=$false
      $view.Add($state.chkNoPreAuth); $y+=1

      $state.chkUseDES = [Terminal.Gui.CheckBox]::new("Use DES encryption types"); $state.chkUseDES.X=4; $state.chkUseDES.Y=$y; $state.chkUseDES.Checked=($computer.UseDESKeyOnly??$false); $state.chkUseDES.Enabled=$false
      $view.Add($state.chkUseDES); $y+=2

      ## Kerberos
      $lbl = [Terminal.Gui.Label]::new("Kerberos Encryption"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($computer.KerberosEncryptionType) {
        $kerbTypes = $computer.KerberosEncryptionType -join ', '
        $lbl = [Terminal.Gui.Label]::new("Encryption types: " + $kerbTypes); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.'msDS-SupportedEncryptionTypes') {
        $lbl = [Terminal.Gui.Label]::new("Supported encryption: " + $computer.'msDS-SupportedEncryptionTypes'); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      ## SID
      $y+=1
      $lbl = [Terminal.Gui.Label]::new("Identifiers"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $sid = $computer.SID ?? $computer.objectSid
      if ($sid) {
        $lbl = [Terminal.Gui.Label]::new("SID:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $state.txtSID = [Terminal.Gui.Label]::new($sid.ToString()); $state.txtSID.X=25; $state.txtSID.Y=$y
        $view.Add($state.txtSID); $y+=1
      }

      if ($computer.ObjectGUID) {
        $lbl = [Terminal.Gui.Label]::new("Object GUID:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $state.txtGUID = [Terminal.Gui.Label]::new($computer.ObjectGUID.ToString()); $state.txtGUID.X=25; $state.txtGUID.Y=$y
        $view.Add($state.txtGUID); $y+=1
      }
    }
  }

  # Member Of Tab
  $memberOfTab = @{
    Name = "Member Of"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Group Memberships:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.lstGroups = [Terminal.Gui.ListView]::new()
      $state.lstGroups.X = 2
      $state.lstGroups.Y = $y
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

      ## NOTE: Computer objects use Show-EditGroupMembershipDialog same as users
      ## We create a pseudo-user object for compatibility
      $pseudoUser = [PSCustomObject]@{
        Name = $computer.Name
        SamAccountName = $computer.SamAccountName
        MemberOf = $computer.MemberOf
        Groups = $state.groupList
      }

      ## Add to Group button
      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2
      $btnAdd.Y = 23
      $btnAdd.add_Clicked({
        Show-EditGroupMembershipDialog -User $pseudoUser -OnUpdate {
          Debug-Log ": Refreshing computer group list after add" -Type "Info"

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
          Debug-Log ": Computer group list refreshed, now showing $($refreshedGroups.Count) groups" -Type "Success"
        }
      }.GetNewClosure())
      $view.Add($btnAdd)

      ## Remove from Group button
      $btnRemove = [Terminal.Gui.Button]::new("Remove from Group")
      $btnRemove.X = 22
      $btnRemove.Y = 23
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
                # Demo mode - update computer's MemberOf
                $computer.MemberOf = $computer.MemberOf | Where-Object { $_ -notmatch "CN=$selectedGroup," }
                Debug-Log ": Removed $($computer.Name) from group $selectedGroup (demo mode)" -Type "Success"
              } else {
                # Production mode
                Remove-ADGroupMember -Identity $selectedGroup -Members $computer.SamAccountName -Confirm:$false
                Debug-Log ": Removed $($computer.Name) from group $selectedGroup" -Type "Success"
              }

              # Refresh the list
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
              Debug-Log ": Failed to remove from group: $($_.Exception.Message)" -Type "Error"
            }
          }
        } else {
          Show-Modal "Info" "Please select a group to remove"
        }
      }.GetNewClosure())
      $view.Add($btnRemove)
    }
  }

  # SPNs Tab
  $spnTab = @{
    Name = "SPNs"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Service Principal Names:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.lstSPNs = [Terminal.Gui.ListView]::new()
      $state.lstSPNs.X = 2
      $state.lstSPNs.Y = $y
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

  # LAPS Tab
  $lapsTab = @{
    Name = "LAPS"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Local Administrator Password Solution"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      ## Check for LAPS in production mode
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
          Debug-Log ": Failed to retrieve LAPS info: $_" -Type "Warn"
        }
      } elseif ($computer.'ms-Mcs-AdmPwd') {
        ## Demo mode with LAPS data
        $lapsEnabled = $true
        $lapsPassword = $computer.'ms-Mcs-AdmPwd'
        $lapsExpiry = if ($computer.'ms-Mcs-AdmPwdExpirationTime') {
          [DateTime]::FromFileTime($computer.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
        } else {
          "Unknown"
        }
      }

      if ($lapsEnabled) {
        $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✓ Enabled"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2
        $lbl = [Terminal.Gui.Label]::new("Administrator Password:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $state.txtPassword = [Terminal.Gui.TextField]::new($lapsPassword); $state.txtPassword.X=4; $state.txtPassword.Y=$y; $state.txtPassword.Width=70; $state.txtPassword.ReadOnly=$true
        $view.Add($state.txtPassword); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Password Expires:"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
        $state.txtExpiry = [Terminal.Gui.Label]::new($lapsExpiry); $state.txtExpiry.X=25; $state.txtExpiry.Y=$y
        $view.Add($state.txtExpiry); $y+=2

        $lbl = [Terminal.Gui.Label]::new("⚠ This password provides full local administrator access"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $lbl = [Terminal.Gui.Label]::new("  Handle with appropriate security controls"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      } else {
        $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✗ Not Enabled"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2
        $lbl = [Terminal.Gui.Label]::new("This computer does not have LAPS configured or you"); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $lbl = [Terminal.Gui.Label]::new("do not have permissions to view the password."); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      }
    }
  }

  # Advanced Tab
  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Distinguished Name:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=1
      $state.txtDN = [Terminal.Gui.TextField]::new($computer.DistinguishedName ?? ""); $state.txtDN.X=2; $state.txtDN.Y=$y; $state.txtDN.Width=90; $state.txtDN.ReadOnly=$true
      $view.Add($state.txtDN); $y+=2

      $lbl = [Terminal.Gui.Label]::new("Canonical Name:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=1
      $state.txtCN = [Terminal.Gui.TextField]::new($computer.CanonicalName ?? ""); $state.txtCN.X=2; $state.txtCN.Y=$y; $state.txtCN.Width=90; $state.txtCN.ReadOnly=$true
      $view.Add($state.txtCN); $y+=2

      ## Timestamps
      $lbl = [Terminal.Gui.Label]::new("Timestamps"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $created = if ($computer.Created -or $computer.whenCreated) {
        ($computer.Created ?? $computer.whenCreated).ToString('yyyy-MM-dd HH:mm:ss')
      } else { "Unknown" }
      $lbl = [Terminal.Gui.Label]::new("Created: " + $created); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

      $modified = if ($computer.Modified -or $computer.whenChanged) {
        ($computer.Modified ?? $computer.whenChanged).ToString('yyyy-MM-dd HH:mm:ss')
      } else { "Unknown" }
      $lbl = [Terminal.Gui.Label]::new("Modified: " + $modified); $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

      if ($computer.LastBootUpTime) {
        $lbl = [Terminal.Gui.Label]::new("Last Boot: " + $computer.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      ## USN
      $y+=1
      $lbl = [Terminal.Gui.Label]::new("Update Sequence Numbers"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($computer.uSNCreated) {
        $lbl = [Terminal.Gui.Label]::new("USN Created: " + $computer.uSNCreated); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.uSNChanged) {
        $lbl = [Terminal.Gui.Label]::new("USN Changed: " + $computer.uSNChanged); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      ## Additional properties
      $y+=1
      $lbl = [Terminal.Gui.Label]::new("Additional Properties"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      if ($computer.isCriticalSystemObject) {
        $lbl = [Terminal.Gui.Label]::new("✓ Critical System Object"); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.ProtectedFromAccidentalDeletion) {
        $lbl = [Terminal.Gui.Label]::new("✓ Protected from Accidental Deletion"); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }

      if ($computer.PrimaryGroup) {
        $lbl = [Terminal.Gui.Label]::new("Primary Group: " + $computer.PrimaryGroup); $lbl.X=4; $lbl.Y=$y
        $view.Add($lbl); $y+=1
      }
    }
  }

## Add in a search properties tab
# ========================= Search / Lookup (Computer) Tab =========================
$searchComputerTab = @{
  Name = "Search/Lookup"
  Builder = {
    param($view, $computer, $state)

    $y = 1

    # ---- Domain ----
    $lblSearchDomain = [Terminal.Gui.Label]::new("Domain:")
    $lblSearchDomain.X = 2
    $lblSearchDomain.Y = $y
    $view.Add($lblSearchDomain)

    $state.txtSearchDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain ?? "")
    $state.txtSearchDomain.X = 15
    $state.txtSearchDomain.Y = $y
    $state.txtSearchDomain.Width = 30
    $view.Add($state.txtSearchDomain)
    $y += 2

    # ---- Computer Name ----
    $lblSearchName = [Terminal.Gui.Label]::new("Computer:")
    $lblSearchName.X = 2
    $lblSearchName.Y = $y
    $view.Add($lblSearchName)

    $state.txtSearchUser = [Terminal.Gui.TextField]::new($computer.Name ?? "")
    $state.txtSearchUser.X = 15
    $state.txtSearchUser.Y = $y
    $state.txtSearchUser.Width = 30
    $view.Add($state.txtSearchUser)
    $y += 2

    # ---- Search Type ----
    $lblSearchType = [Terminal.Gui.Label]::new("Type:")
    $lblSearchType.X = 2
    $lblSearchType.Y = $y
    $view.Add($lblSearchType)

    $state.cmbSearchType = [Terminal.Gui.ComboBox]::new()
    $state.cmbSearchType.X = 15
    $state.cmbSearchType.Y = $y
    $state.cmbSearchType.Width = 20
    $state.cmbSearchType.SetSource(@("Computer","Group","OU"))
    $state.cmbSearchType.SelectedItem = 0
    $view.Add($state.cmbSearchType)
    $y += 2

    # ---- Filter Results ----
    $lblSearchFilter = [Terminal.Gui.Label]::new("Filter Results:")
    $lblSearchFilter.X = 48
    $lblSearchFilter.Y = 1
    $view.Add($lblSearchFilter)

    $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
    $state.txtSearchFilter.X = 62
    $state.txtSearchFilter.Y = 1
    $state.txtSearchFilter.Width = 20
    $view.Add($state.txtSearchFilter)

    $state.txtSearchFilter.add_TextChanged({
      if ($state.currentSearchOutputLines) {
        $search = $state.txtSearchFilter.Text.ToString().Trim()
        if ($search) {
          $state.txtSearchOutput.Text =
            ($state.currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
        }
        else {
          $state.txtSearchOutput.Text = $state.currentSearchOutputLines -join "`n"
        }
      }
    }.GetNewClosure())

    # ---- Results ----
    $lblSearchResult = [Terminal.Gui.Label]::new("Results:")
    $lblSearchResult.X = 2
    $lblSearchResult.Y = $y
    $view.Add($lblSearchResult)
    $y += 1

    $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
    $state.txtSearchOutput.X = 2
    $state.txtSearchOutput.Y = $y
    $state.txtSearchOutput.Width  = [Terminal.Gui.Dim]::Fill(2)
    $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(4)
    $state.txtSearchOutput.ReadOnly = $true
    $state.txtSearchOutput.WordWrap = $false
    $view.Add($state.txtSearchOutput)

    # ---- Computer State ----
    $state.chkSearchLocked = [Terminal.Gui.CheckBox]::new("Enabled")
    $state.chkSearchLocked.X = 2
    $state.chkSearchLocked.Y = [Terminal.Gui.Pos]::Bottom($state.txtSearchOutput) + 1
    $state.chkSearchLocked.CanFocus = $true
    $view.Add($state.chkSearchLocked)

    $state.chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Disabled")
    $state.chkSearchDisabled.X = 2
    $state.chkSearchDisabled.Y = [Terminal.Gui.Pos]::Bottom($state.chkSearchLocked) + 1
    $state.chkSearchDisabled.CanFocus = $true
    $view.Add($state.chkSearchDisabled)

    # ---- Search Button (UI only for now) ----
    $btnDoSearch = [Terminal.Gui.Button]::new("Search")
    $btnDoSearch.X = 48
    $btnDoSearch.Y = 3
    $view.Add($btnDoSearch)

    # ---- Auto-populate computer properties ----
    [Terminal.Gui.Application]::MainLoop.Invoke({
      if ($computer) {

        $lines = @()
        $computer.PSObject.Properties | ForEach-Object {
          $value = if ($_.Value -is [array]) {
            $_.Value -join ', '
          }
          elseif ($null -eq $_.Value) {
            ''
          }
          else {
            $_.Value.ToString()
          }

          $lines += "$($_.Name.PadRight(25)): $value"
        }

        $state.txtSearchOutput.Text = $lines -join "`n"
        $state.currentSearchOutputLines = $lines

        $state.chkSearchLocked.Checked   = [bool]$computer.Enabled
        $state.chkSearchDisabled.Checked = -not [bool]$computer.Enabled
      }
    }.GetNewClosure())
  }
}

  ## ==================== Apply Logic ====================
  # Computer properties are mostly read-only, but we can update Description and Location
  $applyLogic = {
    param($computer, $state)

    Debug-Log ": Apply clicked - saving computer changes" -Type "Info"

    try {
      $changesMade = $false

      # Check Description change
      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $computer.Description) {
          if ($Script:DemoMode) {
            $computer.Description = $newDescription
            $changesMade = $true
          } else {
            Set-ADComputer -Identity $computer.SamAccountName -Description $newDescription -ErrorAction Stop
            $computer.Description = $newDescription
            $changesMade = $true
          }
        }
      }

      # Check Location change
      if ($state.txtLocation) {
        $newLocation = $state.txtLocation.Text.ToString().Trim()
        if ($newLocation -ne $computer.Location) {
          if ($Script:DemoMode) {
            $computer.Location = $newLocation
            $changesMade = $true
          } else {
            Set-ADComputer -Identity $computer.SamAccountName -Location $newLocation -ErrorAction Stop
            $computer.Location = $newLocation
            $changesMade = $true
          }
        }
      }

      if ($changesMade) {
        Show-Modal "Success" "Computer changes applied successfully"
      } else {
        Show-Modal "Info" "No changes to apply"
      }

    } catch {
      Show-Modal "Error" "Failed to apply computer changes:`n$($_.Exception.Message)"
      Debug-Log ": Apply failed: $($_.Exception.Message)" -Type "Error"
    }
  }

  ## ==================== Create and Show Dialog ====================
  $tabs = @($generalTab, $accountTab, $securityTab, $memberOfTab, $spnTab, $lapsTab, $advancedTab, $searchComputerTab)

  New-PropertiesDialog -Title "Computer Properties - $($computer.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $computer -OnApply $applyLogic
}

# -------------------------[ Context Menu Handler }-------------------------
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
      "Reset Password"    { Show-ResetPasswordDialog -userName $obj.Name                             }
      "Disable Account"   { Toggle-UserAccount -userName $obj.Name -disable $true                    }
      "Enable Account"    { Toggle-UserAccount -userName $obj.Name -disable $false                   }
      "Unlock Account"    { Unlock-UserAccount -userName $obj.Name                                   }
      "Disable"           { Toggle-ComputerAccount -computerName $obj.Name -disable $true            }
      "Enable"            { Toggle-ComputerAccount -computerName $obj.Name -disable $false           }
      "Move to OU..."     { Invoke-ObjectOperation -Objects @($selectedNode.Tag) -Operation 'Move'   }
      "Delete"            { Invoke-ObjectOperation -Objects @($selectedNode.Tag) -Operation 'Delete' }
      "Add Member..."     { Show-EditGroupMembershipDialog -groupName $obj.Name                      }
      "Remove Member..."  { Show-EditGroupMembershipDialog -groupName $obj.Name                      }
      "Check Replication" { Check-DCReplication -dcName $obj.Name                                    }
      "Refresh"           { Refresh-TreeData                                                         }
    }
  }
})

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
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

  ## ---- Containers / OUs (no backing object) ----
  #if (-not $tag -or -not $tag.Object) {
  #    Debug-Log ": Container / OU selected (Type: $($tag.Type))" -Type "Info"
  #    return
  #}

  ## Get the actual AD object from the Tag
  $obj = $tag.Object

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
      Show-Modal "Domain Controllers" "This is a container for Domain Controllers in this domain.`n`nSelect an individual DC to view its properties."
      return
    }
    'computer' {
      Debug-Log ": COMPUTER object selected: $($obj.Name)" -Type "Info"
      Show-ComputerPropertiesDialog -computerName $obj.Name  # ← Changed to -computerName and .Name
      return
    }
    'ou' {
      Debug-Log ": Showing OU properties for $($obj.Name)" -Type "Info"
      Show-OUPropertiesDialog -ou $obj
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
      Refresh-TreeData
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
## Never called - confirm if it's still needed or is dead code for removal
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

## Add keyboard shortcuts for selection
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
  if (-not $Script:DemoMode) { Load-DomainData -domain $Script:CurrentDomain }
  Build-Tree -domain $Script:CurrentDomain
  Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel

  ## Clear selection
  $Script:SelectedObjects = @()
  $Script:SelectionMode = $false
  Update-SelectionPanel -panel $selectionPanel
  }
}

## -------------------------{} Bulk Add to Group }-------------------------
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
  })
  $dlg.AddButton($btnAdd)

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)

  [Terminal.Gui.Application]::Run($dlg)
}

## Add keyboard shortcuts for selection
## Never callled - confirm if it needs ot be, or is dead code for removal
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
      if ($errors.Count -gt 0 -and $errors.Count -le 5) {
        $msg += "`n`nErrors:`n" + ($errors -join "`n")
      }
    }

    Show-Modal "Bulk Action Complete" "$msg"
    ## Refresh tree
    if (-not $Script:DemoMode) {
      Load-DomainData -domain $Script:CurrentDomain
    }
    Build-Tree -domain $Script:CurrentDomain
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel

    ## Clear selection
    $Script:SelectedObjects = @()
    $Script:SelectionMode = $false
    Update-SelectionPanel -panel $selectionPanel
  }
}

# -------------------------{} Bulk Add to Group }-------------------------
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
  })
  $dlg.AddButton($btnAdd)

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

## -------------------------{ Selection Panel }-------------------------
function Create-SelectionPanel {
  param(
    [int]$panelWidth = 30,
    [int]$panelHeight = 15
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

## ===================================================={ ADD / REMOVE GROUP MEMBERS AKA EDIT }====================================================
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
    $currentGroups = $User.MemberOf | ForEach-Object {
    if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
    }
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
    if ($toAdd.Count -gt 0) {
      $changeMsg += "Add to $($toAdd.Count) group(s):`n  " + ($toAdd -join "`n  ") + "`n`n"
    }
    if ($toRemove.Count -gt 0) {
      $changeMsg += "Remove from $($toRemove.Count) group(s):`n  " + ($toRemove -join "`n  ")
    }

    $confirm = Show-Modal "Confirm Changes" "Apply these changes for $($User.Name)?`n`n$changeMsg" -YesNo

    if ($confirm -ne 0) {  # Not "Yes"
      return
    }

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
  })
  $dlg.Add($btnApply)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.X = [Terminal.Gui.Pos]::Right($btnApply) + 2
  $btnCancel.Y = [Terminal.Gui.Pos]::Bottom($scrollView) + 1
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.Add($btnCancel)

  ## Run the dialog
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

function Show-OUPropertiesDialog {
  param($ou)
  ## ---------------------- Safety Checks ----------------------
  Debug-Log ": Showing OU properties dialog" -Type "Info"
  Debug-Log ": OU object type: $($ou.GetType().Name)" -Type "Info"
  Debug-Log ": OU properties: $($ou | ConvertTo-Json -Depth 1)" -Type "Info"
  if (-not $ou) {
    Debug-Log ": OU object is null" -Type "Warn"
    return
  }
  ## Try to get the name from wherever it actually is
  $ouName = $ou.Name ?? $ou.Text ?? "Unknown"
  Debug-Log ": OU name resolved to: $ouName" -Type "Info"
  ## If not found, create a basic one (for existing OUs from user data)
  if (-not $ou.Name) {
    $ou = @{
      Name = $ouName
      Path = ""
      Description = ""
    }
  }
  ## ==================== Tab Definition ====================

  ## ---------------------- General Tab ----------------------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $ou, $state)
      $y = 1
      ## OU Name (editable for rename)
      $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtName = [Terminal.Gui.TextField]::new($ou.Name ?? "")
      $state.txtName.X=20; $state.txtName.Y=$y; $state.txtName.Width=50
      $view.Add($state.txtName)
      $y+=2
      ## Description
      $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtDesc = [Terminal.Gui.TextField]::new($ou.Description ?? "")
      $state.txtDesc.X=20; $state.txtDesc.Y=$y; $state.txtDesc.Width=50
      $view.Add($state.txtDesc)
      $y+=2
      ## Path (read-only)
      $lbl = [Terminal.Gui.Label]::new("Path:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.txtPath = [Terminal.Gui.TextField]::new($ou.Path ?? "")
      $state.txtPath.X=20; $state.txtPath.Y=$y; $state.txtPath.Width=50
      $state.txtPath.ReadOnly=$true
      $view.Add($state.txtPath)
      $y+=2
      ## Show object count in this OU
      $objectCount = ($Script:Users | Where-Object { $_.OU -contains $ou.Name }).Count
      $lbl = [Terminal.Gui.Label]::new("Contains:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
      $state.lblCount = [Terminal.Gui.Label]::new("$objectCount objects")
      $state.lblCount.X=20; $state.lblCount.Y=$y
      $view.Add($state.lblCount)
    }
  }

  ## ---------------------- Statistics Tab ----------------------
  $statisticsTab = @{
    Name = "Statistics"
    Builder = {
      param($view, $ou, $state)

      ## Calculate statistics
      $ouName = $ou.Name

      ## Count objects in this OU
      $usersInOU = $Script:Users | Where-Object { $_.OU -contains $ouName }
      $enabledUsers = ($usersInOU | Where-Object { $_.Enabled -eq $true }).Count
      $disabledUsers = ($usersInOU | Where-Object { $_.Disabled -eq $true }).Count
      $lockedUsers = ($usersInOU | Where-Object { $_.LockedOut -eq $true }).Count

      $groupsInOU = if ($Script:Groups) {
        $Script:Groups | Where-Object { $_.OU -contains $ouName }
      } else { @() }

      $computersInOU = if ($Script:Computers) {
        $Script:Computers | Where-Object { $_.OU -contains $ouName }
      } else { @() }

      $enabledComputers = ($computersInOU | Where-Object { $_.Enabled -eq $true }).Count
      $disabledComputers = ($computersInOU | Where-Object { $_.Disabled -eq $true }).Count

      ## Nested OUs
      $nestedOUs = if ($Script:rawOUs) {
        $Script:rawOUs | Where-Object { $_.Path -match [regex]::Escape($ouName) -and $_.Name -ne $ouName }
      } else { @() }

      $totalObjects = $usersInOU.Count + $groupsInOU.Count + $computersInOU.Count

      ## Display statistics
      $y = 1

      ## Header
      $lblHeader = [Terminal.Gui.Label]::new("OU: $ouName")
      $lblHeader.X = 2; $lblHeader.Y = $y
      $view.Add($lblHeader)
      $y += 2

      ## Users section
      $lblUsers = [Terminal.Gui.Label]::new("═══ USERS ═══")
      $lblUsers.X = 2; $lblUsers.Y = $y
      $view.Add($lblUsers)
      $y += 1

      $lblUserTotal = [Terminal.Gui.Label]::new("Total Users:       $($usersInOU.Count)")
      $lblUserTotal.X = 4; $lblUserTotal.Y = $y
      $view.Add($lblUserTotal)
      $y += 1

      $lblUserEnabled = [Terminal.Gui.Label]::new("  Enabled:         $enabledUsers")
      $lblUserEnabled.X = 4; $lblUserEnabled.Y = $y
      $view.Add($lblUserEnabled)
      $y += 1

      $lblUserDisabled = [Terminal.Gui.Label]::new("  Disabled:        $disabledUsers")
      $lblUserDisabled.X = 4; $lblUserDisabled.Y = $y
      $view.Add($lblUserDisabled)
      $y += 1

      $lblUserLocked = [Terminal.Gui.Label]::new("  Locked Out:      $lockedUsers")
      $lblUserLocked.X = 4; $lblUserLocked.Y = $y
      $view.Add($lblUserLocked)
      $y += 2

      ## Computers section
      $lblComputers = [Terminal.Gui.Label]::new("═══ COMPUTERS ═══")
      $lblComputers.X = 2; $lblComputers.Y = $y
      $view.Add($lblComputers)
      $y += 1

      $lblCompTotal = [Terminal.Gui.Label]::new("Total Computers:   $($computersInOU.Count)")
      $lblCompTotal.X = 4; $lblCompTotal.Y = $y
      $view.Add($lblCompTotal)
      $y += 1

      $lblCompEnabled = [Terminal.Gui.Label]::new("  Enabled:         $enabledComputers")
      $lblCompEnabled.X = 4; $lblCompEnabled.Y = $y
      $view.Add($lblCompEnabled)
      $y += 1

      $lblCompDisabled = [Terminal.Gui.Label]::new("  Disabled:        $disabledComputers")
      $lblCompDisabled.X = 4; $lblCompDisabled.Y = $y
      $view.Add($lblCompDisabled)
      $y += 2

      ## Groups section
      $lblGroups = [Terminal.Gui.Label]::new("═══ GROUPS ═══")
      $lblGroups.X = 2; $lblGroups.Y = $y
      $view.Add($lblGroups)
      $y += 1

      $lblGroupTotal = [Terminal.Gui.Label]::new("Total Groups:      $($groupsInOU.Count)")
      $lblGroupTotal.X = 4; $lblGroupTotal.Y = $y
      $view.Add($lblGroupTotal)
      $y += 2

      ## Structure section
      $lblStructure = [Terminal.Gui.Label]::new("═══ STRUCTURE ═══")
      $lblStructure.X = 2; $lblStructure.Y = $y
      $view.Add($lblStructure)
      $y += 1

      $lblNestedOUs = [Terminal.Gui.Label]::new("Nested OUs:        $($nestedOUs.Count)")
      $lblNestedOUs.X = 4; $lblNestedOUs.Y = $y
      $view.Add($lblNestedOUs)
      $y += 2

      ## Total
      $lblTotal = [Terminal.Gui.Label]::new("═══════════════════")
      $lblTotal.X = 2; $lblTotal.Y = $y
      $view.Add($lblTotal)
      $y += 1

      $lblTotalObjects = [Terminal.Gui.Label]::new("TOTAL OBJECTS:     $totalObjects")
      $lblTotalObjects.X = 2; $lblTotalObjects.Y = $y
      $view.Add($lblTotalObjects)
      $y += 2

      ## Export button
      $btnExport = [Terminal.Gui.Button]::new("Export Statistics")
      $btnExport.X = 2
      $btnExport.Y = $y
      $btnExport.add_Clicked({
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "ou_stats_${ouName}_$timestamp.csv"

        try {
          $stats = [PSCustomObject]@{
            OU = $ouName
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            TotalUsers = $usersInOU.Count
            EnabledUsers = $enabledUsers
            DisabledUsers = $disabledUsers
            LockedUsers = $lockedUsers
            TotalComputers = $computersInOU.Count
            EnabledComputers = $enabledComputers
            DisabledComputers = $disabledComputers
            TotalGroups = $groupsInOU.Count
            NestedOUs = $nestedOUs.Count
            TotalObjects = $totalObjects
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
    ## Call existing Apply-ObjectChanges function
    Apply-ObjectChanges -ObjectType 'OU' -Object $ou -State $state
  }
  ## ==================== Create and Show Dialog ====================
  $tabs = @($generalTab, $statisticsTab)
  Debug-Log ": Show-OUPropertiesDialog running" -Type "Info"
  New-PropertiesDialog -Title "OU Properties - $ouName" -Width 80 -Height 28 -Tabs $tabs -Data $ou -OnApply $applyLogic
  Debug-Log ": Show-OUPropertiesDialog completed" -Type "Info"
}

#################################### Program Launch Begins Here ####################################

## ===== STEP 1: Environment & Logging =====

## Don't let users do stupid stuff, either unintentionally or willingly
if ($DemoMode -and $PSBoundParameters.ContainsKey('Domain')) {
  Debug-Log "Invalid startup: -Domain cannot be used with -DemoMode" -Type "Error"
  return
}

## Echo basic info for debugging
Debug-Log "DemoMode: $DemoMode" -Type "info"
Debug-Log "Logging: $Logging" -Type "Info"
Debug-Log "LogFile: $LogFile" -Type "Info"

## Initialize logging if requested
if ($Logging -or $LogFile) {
  Debug-Log "Logging condition TRUE" "Debug"

  if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "dsa_tui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Debug-Log "Auto-generated LogFile: $LogFile" -Type "Debug"
  }

  ## Convert to absolute path if needed
  if (-not [System.IO.Path]::IsPathRooted($LogFile)) {
    $LogFile = Join-Path (Get-Location).Path $LogFile
  }

  $Script:Logging = $true
  $Script:LogFile = $LogFile

  ## Dummy status object for UI
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

## Set globals for demo mode and theme
$Script:DemoMode  = $DemoMode
$Script:ThemeMode = $Theme

## ===== STEP 2: Module Checks & Terminal.Gui =====
Debug-Log "Performing pre-flight module checks..." -Type "Info"

## Required module: Terminal.Gui via ConsoleGuiTools
$requiredOK = Test-RequiredModule -Name "Microsoft.PowerShell.ConsoleGuiTools"
if (-not $requiredOK) {
  Debug-Log "Missing required module Microsoft.PowerShell.ConsoleGuiTools. Exiting." -Type "Error"
  exit
}

## Optional module: ActiveDirectory
$adAvailable = Test-RequiredModule -Name "ActiveDirectory" -Optional
if (-not $adAvailable) {
  Debug-Log "ActiveDirectory module missing. Falling back to DEMO mode..." -Type "Warn"
  $Script:DemoMode = $true
}

## Optional modules
$Script:HasPSWriteColor = Test-RequiredModule -Name "PSWriteColor" -Optional
$null = Test-RequiredModule -Name "Terminal-Icons" -Optional

$Script:UseIcons = $false
if (Get-Module -Name Terminal-Icons -ErrorAction SilentlyContinue) {
  try { Write-Host '' -NoNewline; $Script:UseIcons = $true } catch {}
}

## Ensure Terminal.Gui.dll is loaded
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

## Colours and demo mode
$Script:DemoMode = $DemoMode
$Script:ThemeMode = $Theme

## Spinner setup
#$Script:SpinnerTimer  = $null
#$Script:SpinnerActive = $false

## Tab & layout placeholders
$Script:LayoutInProgress = $false
$Script:TabRows          = @()
$Script:AllTabs          = @()
$Script:ActiveTab        = $null
$Script:TabRowHeight     = 1

## ===== STEP 5: Initialize Terminal.Gui UI
## Check if date is special and use emoji accordingly
Initialize-DirectoryEmoji
$windowTitle = "$($Script:ProjectName) $($Script:DirectoryEmoji) Active Directory $BuildVersion Codename: $($Script:FruitName)"

## Initialize UI framework FIRST
$uiComponents = Initialize-UIFramework -Theme $Theme -Title $windowTitle

## Extract components
$top = $uiComponents.Top
$win = $uiComponents.Window
$Script:themeData = $uiComponents.Theme

Debug-Log "UI Framework ready - window visible to user" -Type "Success"
Dump-ColourScheme

## Forest/domain structure
if ($Script:DemoMode) {
  Debug-Log "DemoMode enabled: creating demo forest structure..." -Type "Info"
  $Script:ForestName    = "jukebox.example"
  $Script:RootDomain    = "example.com"
  $Script:Domains       = @('example.com', 'example.net', 'example.org')
  ## causes a white screen, i.e.e initialisation is busted
  #$Script:Domains       = $Script:rawDCs.Domain | Where-Object { $_ } | Sort-Object -Unique
  ##$Script:Sites         = $Script:rawDCs.Site | Sort-Object -Unique
  ## TODO This really ought to be enumerated, as hard-coded really messes produciton up
  $Script:Sites = @( 'ABR', 'BIR', 'BON', 'BRD', 'BRL', 'BRK', 'CLY', 'CPH', 'DUN', 'EDI', 'FAX', 'GLA', 'KGE', 'KRS', 'LIV', 'LND', 'MCR', 'MIA', 'MUC', 'MUN', 'NEW', 'ODE', 'VIE' )

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

    $Script:ForestName    = $forest.Name.Split('.')[0].ToUpper()
    $Script:RootDomain    = $forest.RootDomain
    $Script:Domains       = $forest.Domains
    $Script:Sites         = $forest.Sites | ForEach-Object { $_.Name }
    $Script:CurrentDomain = if ($Domain) { $Domain } else { $Script:RootDomain }
  } catch {
    Debug-Log ("Failed to query AD domain/forest: $_") -Type "Error"
    Debug-Log ("Falling back to minimal domain info.") -Type "Warn"
    $Script:ForestName    = "DOMAIN"
    $Script:RootDomain    = if ($Domain) { $Domain } else { $env:USERDNSDOMAIN }
    $Script:Domains       = @($Script:RootDomain)
    $Script:Sites         = @()
    $Script:CurrentDomain = $Script:RootDomain
  }
}

$Script:Domain = $Script:CurrentDomain  ## Compatibility

## Initialize object arrays
$Script:CurrentDC       = $null
$Script:Users           = @()
$Script:Groups          = @()
$Script:DCs             = @()
$Script:ADObjects       = @()
$Script:SelectedObjects = @()
$Script:SelectionMode   = $false

## Global Search filters
$Script:FilterOptions = @{
  ShowDisabledUsers   = $true
  ShowEnabledUsers    = $true
  ShowLockedUsers     = $true
  ShowGroups          = $true
  ShowDCs             = $true
  ShowComputers       = $true
  ShowOUs             = $true
  NameFilter          = ""
  SortBy              = "Name"
  SortDescending      = $false
}

# -------------------------{ Status Bar }-------------------------
## Create status bar FIRST so updates are visible
Debug-Log "Creating status bar..." -Type "Info"
$statusBar = Set-StatusBar -Initialize -ThemeData $Script:themeData
$top.Add($statusBar)

# -------------------------{ Load Domain Data }-------------------------
## Now load data with visible status updates
Debug-Log "Loading domain data for $($Script:CurrentDomain)..." -Type "Info"
Set-StatusBar "Loading domain data for $($Script:CurrentDomain)..." -Spinner
Load-DomainData -domain $Script:CurrentDomain
Set-StatusBar "Ready" -Final

Debug-Log ("POST-LOAD: Users=$($Script:Users.Count), Objects=$($Script:ADObjects.Count), DCs=$($Script:DCs.Count)") -Type "Info"
Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Info"

## ===== STEP 6: Build UI Components =====

## -------------------------{ Menu }-------------------------
Debug-Log ": Creating main menu..." -Type "Info"
$menu = Build-MainMenu
$top.Add($menu)

## -------------------------{ Filter Panel (right, top) }-------------------------
Debug-Log ": Creating filter panel..." -Type "Info"
$filterPanel = Create-FilterPanel
if (-not ($filterPanel -is [Terminal.Gui.View])) {
  $filterPanel = [Terminal.Gui.FrameView]::new("Filters")
}
$filterPanel.Width  = 40
$filterPanel.Height = 27  # Matches Create-FilterPanel height
$filterPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$filterPanel.Y = 0
$win.Add($filterPanel)

## -------------------------{ Selected Objects Panel (right, below filters) }-------------------------
Debug-Log ": Creating selection panel..." -Type "Info"
$selectedObjectsPanel = Create-SelectionPanel
if (-not ($selectedObjectsPanel -is [Terminal.Gui.View])) {
  $selectedObjectsPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
}
$selectedObjectsPanel.Width  = 40
$selectedObjectsPanel.Height = 10
$selectedObjectsPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$selectedObjectsPanel.Y = 27  # Below filter panel
$win.Add($selectedObjectsPanel)

## -------------------------{ TreeView }-------------------------
Debug-Log ": Initializing TreeView..." -Type "Info"
Set-StatusBar "Building tree view..." -Spinner

## Create a FrameView to hold the tree
$treeFrame = [Terminal.Gui.FrameView]::new("Active Directory Objects")
$treeFrame.X = 0
$treeFrame.Y = 0
$treeFrame.Width = [Terminal.Gui.Dim]::Fill(42)  # Leave room for right panels
$treeFrame.Height = [Terminal.Gui.Dim]::Fill()

## Create and configure the tree
$Script:tree = [Terminal.Gui.TreeView]::new()
$Script:tree.X = 0
$Script:tree.Y = 0
$Script:tree.Width = [Terminal.Gui.Dim]::Fill()
$Script:tree.Height = [Terminal.Gui.Dim]::Fill()

## Add right-click context menu handler
$Script:tree.add_MouseClick({
  param($senders, $arguments)

  ## Check if it's a right-click (Button3)
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

    ## Get the object and its type
    $obj = $tag.Object
    $objType = $tag.Type
    Debug-Log ": Right-click on object: $($obj.Name), Type: $objType" -Type "Info"

    ## Build context menu based on type
    $menuItems = Build-ContextMenuItems -ObjectType $objType -Object $obj

    ## Show context menu
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

## Add tree to frame, then frame to window
$treeFrame.Add($Script:tree)
$win.Add($treeFrame)
Debug-Log ": TreeView created and added to window successfully" -Type "Success"
Set-StatusBar "Ready" -Final

## -------------------------{ Global Key Handlers }-------------------------
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

## -------------------------{ Debug View Tree (if debug mode enabled) }-------------------------
if ($DebugMode -or $Logging) {
  Debug-Log "=== FULL VIEW TREE DUMP ===" -Type "Info"
  Debug-DumpViewTree -View $top
  Debug-Log "=== END VIEW TREE DUMP ===" -Type "Info"
}

## -------------------------{ Run Application }-------------------------
Debug-Log ": Starting Terminal.Gui main loop..." -Type "Success"
[Terminal.Gui.Application]::Run($top)

## -------------------------{ Cleanup }-------------------------
Debug-Log ": Application stopped, cleaning up..." -Type "Info"
Set-StatusBar "Shutting down"
[Terminal.Gui.Application]::Shutdown()
Debug-Log "Application shut down cleanly" -Type "Success"

## He fights for the users...
Debug-Log "End of line..." -Type "Info"
