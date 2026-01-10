# DSA-TUI Blåbær — Active Directory TUI Tool

## Historical Build Notes and Change Log (Newest → Oldest)

---

### 3.0.0.31 (Bug fix)
- Users are now enumerated properly and shown in real AD.
- LAPS password modal now shows correctly.
- Group, OU, Computer and DC objects now all have a Properties tab.
- AD tools added into the AD Health modal.
- AD Health reworked to be more feature rich.
- Change domain bug fixed (first refresh crash resolved).
- Change domain now falls back to previous domain on failure.

---

### 3.0.0.29 (Cleanup and reflow many areas of code)
- Massive consolidation and feature expansion session. Reduced code duplication with unified functions, added bulk
  operations and auditing features.
- **Manage-FilterStatusLabel**
  - Unified Create-FilterStatusLabel + Update-FilterStatusLabel
  - Single function with `-Action 'Create'/'Update'` parameter
  - Supports both standalone and in-panel creation with `-InPanel` switch
  - Replaced: Create-FilterStatusLabel, Update-FilterStatusLabel
- **Manage-Spinner**
  - Unified `Start-Spinner` + `Stop-Spinner`
  - Single function with `-Action 'Start'/'Stop'` parameter
  - Message parameter validation when starting
  - Replaced: Start-Spinner, Stop-Spinner
- **Manage-Selection**
  - Unified Select-AllObjects and Deselect-AllObjects
  - Single function with `-Action 'SelectAll'/'DeselectAll'` parameter
- **Get-Theme**
  - Combined Get-Theme + Dump-ColourScheme
  - Added `-Dump` switch for debugging theme colours
  - All 19 themes preserved
- **Apply-ObjectChanges**
  - Unified User/Group/OU/Computer apply logic
- **Invoke-ObjectOperation**
  - Unified delete/move/bulk operations
- **AD Health refactor**
  - Refactored monolithic Get-ADHealth into modular checks
  - Renamed Get-ADHealth → Check-ADHealth
- Added AD Tools modal.
- Added bulk account enable/disable.
- Added bulk attribute editor.
- Added stale account finder.
- Added AD object cloning.
- Added demo data import/export.
- Introduced New-PropertiesDialog.

---

### 2.3.8.21 (Bug fix and code consolidation)
- Fixed Theme selector and background colour issues.
- Improved Show-Modal Yes/No handling.
- Added more themes.
- Introduced New-PropertiesDialog.
- Rewrote all object property dialogs to use it.
- Added expired LAPS and stale device demo scenarios.
- Reduced redundant functions.
- Added CSV import/export tooling (no validation).

---

### 2.3.8.0 (Code reflow)
- Nerd font icons used when available.
- Removed Get-CleanObjectInfo.
- Simplified object handling and icon detection.
- Improved startup repaint logic.

---

### 2.3.7.2 (Bug fix and cleanup)
- Fixed leftover `$Global` usage.
- Fixed group membership modal.
- Load terminal icons module if present.
- Advise about Nerd Fonts.
- Added non-computer devices.
- Added Altered Images and Eurythmics.
- Added UPNs and SamAccountNames.
- Restored search/lookup tab for users.

---

### 2.3.7.1 (Bug fix)
- Added Initialize-DirectoryEmoji (date-based emojis).

---

### 2.3.7.0 (Logic, order of operations, and plumbing fixed)
- Faux multi-row tabs implemented.
- Standardised Debug-Log output.
- Added TWA Airlines theme.
- Menu creation moved into a function.
- Reworked startup phases.
- Added runtime tree debugging.
- Improved domain switching UX.
- Added LAPS search modal.
- Improved User and Computer properties with tabs.
- Unified status bar logic.
- Enabled copying, pasting and exporting LDAP output.
- Reworked group membership logic.

---

### 2.1.5.2 (Computers, refresh debug, new bands, status bar++)
- Added Computer objects.
- Added Computer properties.
- Enhanced refresh debugging.
- Improved search modal with multi-domain support.
- Enhanced status bar with progress indicators.
- Added new demo bands and locations.
- RFC-compliant example domains used.

---

### 2.0.0.1 (Bug fix release)
- Fixed window fill layout leaving space for status bar.
- Debug log optional file output.
- Documented right-click context menu issues (WIP).
- Increased text field lengths.
- Improved search reactivity.
- Status indicators for locked and disabled accounts.

---

### 2.0.0.0 (Multi-domain support with skittles mode)
- Multi-domain support added.
- Demo-mode refresh crash fixed.
- Major code cleanup.
- Added multiple new themes.
- Two-column theme selector.

---

### 1.8.6.0 (Colour my life)
- Added DSB Danish State Railways and Pan American Airlines themes.

---

### 1.8.4.0 (Refactoring)
- Large refactor and cleanup.
- Added OU, DC and Group editing support.
- Reduced dialog duplication using Show-Modal.
- Demo-mode checks consolidated.

---

### 1.8.3.0 (More cowbell)
- Added missing and former band members.
- Added Get-CleanObjectInfo.
- Fixed properties modal buttons.
- Streamlined `$Script` usage.

---

### 1.8.2.0 (Lumberjack mode)
- Reworked Build-Tree logic.
- Removed duplicate code.
- Added Theme Selector modal.
- Verbose-only debug output.
- Documented why Show-Properties is used.

---

### 1.8.1.0 (Verbosity)
- Added Debug-Log.
- Unified demo and production data models.
- Removed dead demo code.

---

### 1.8.0.0 (Domain Controller information)
- Domain Controller information modal (WIP).
- Replication and synchronisation modal (WIP).
- Fixed TreeView right-click context menu visibility.

---

### 1.7.1.1 (Fixes for regressions)
- Cleaned up demo tree building.
- Fixed Add and Remove button logic.

---

### 1.7.0.1 (Demo data expansion)
- Added devices and office printers.
- Added Rocazino from the Køge office.
- Added multiple Denmark locations.
- Expanded user attributes (e.g. Alan Wilder).

---

### 1.7.0.0 (Demo data and menu fixes)
- Demo data reworked to be more AD-like.
- Dedicated About menu added.
- Shortcuts modal added.

---

### 1.6.9.0 (Window and layout rebuild)
- Clean main window layout.
- Fixed TreeView rendering issues.
- Fixed filter panel behaviour.
- Fixed modal stacking and z-order issues.
- Removed obsolete commented code.

---

### 1.6.3.0 (Password Generator)
- Added password generator modal.
- Clipboard copy support.
- Character set toggles.

---

### 1.5.0.0 (Refresh engine rewrite)
- Major rewrite of refresh and Build-Tree pipeline.
- Searchable attributes added.
- LDAP query optimisation.
- Caching introduced.

---

### 1.4.0.0 (Filter system v1)
- Filter panel added.
- Global FilterOptions introduced.
- Name-based search added.

---

### 1.3.0.0 (Selection and node information)
- Node selection handling added.
- Properties modal introduced.
- Object type icons added (U/G/OU/DC).

---

### 1.2.0.0 (Tree and navigation)
- TreeView implemented.
- Refresh bound to F5.
- Status bar added.

---

### 1.1.0.0 (Initial AD integration)
- LDAP domain bind and query functions.
- Initial Build-Tree prototype.

---

### 1.0.0.0 (Initial experimental)
- First internal test build.
- Basic TUI scaffolding only.
- Placeholder TreeView.
