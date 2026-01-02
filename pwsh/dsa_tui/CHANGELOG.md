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
