## Historical Build Notes and Change Log

### 1.0.0.00  (TUI Conversion from WPF)
- Initial Terminal.Gui conversion from easyDNS v0.2.27 WPF
- Basic DNS zone and record management
- Dashboard view with system info
- Diagnostic tools (Ping, Nslookup, Resolve)

---

### 1.0.1.03  (Feature Enhancement)
- Added DNS record deletion
- Added reverse zone creation
- Import/Export functionality (CSV/JSON/XML)
- Advanced diagnostics (Traceroute, Benchmark, Cache management)
- Auto-refresh system

---

### 1.0.2.06  (Framework Integration)
- Integrated with sophisticated TUI framework
- Added theme support  
  (9 themes: light, dark, matrix, british, panam, dsb, gemstones, class91, scotrail)
- Proper module loading with Terminal.Gui.dll
- Global variables and configuration
- Enhanced modal dialogs
- Debug logging system
- Function key shortcuts
- Status bar integration
- Command line parameters

---

### 1.0.3.00  (Critical Fixes & Zone Reload)
- **CRITICAL FIX:** Removed problematic timer from loading dialog (runspace error)
- **CRITICAL FIX:** Removed `$null` menu separators (Terminal.Gui 1.16 compatibility)
- Added Zone Reload functionality  
  (secondary zones transfer, primary zones reload from file)
- Added File Browser dialog for import/export file selection
- Added DNS Tools Information dialog
- Improved About dialog with codename and credits
- Added Reload button to Forward and Reverse zone views
- Better error handling and logging

---

### 1.0.3.01  (StatusBar & Connection Fixes)
- **CRITICAL FIX:** Simplified StatusBar to avoid Terminal.Gui Key enum CLS compliance errors
- **CRITICAL FIX:** Added null checking in `Get-SafeDnsServerZone`  
  (prevents *"Cannot index into a null array"*)
- **CRITICAL FIX:** Better array handling in zone retrieval
- Changed StatusBar to use simple `Ctrl+Q` quit shortcut instead of multiple F-keys
- Improved error handling in Dashboard statistics
- Updated window title to show connected server name

---

### 1.0.3.02  (Terminal.Gui Key Enum Complete Avoidance)
- **CRITICAL FIX:** Removed ALL StatusItems from StatusBar  
  (Key enum unusable on some systems)
- **CRITICAL FIX:** Changed em-dash to hyphen in window title  
  (encoding/width issues)
- **CRITICAL FIX:** Moved window title update to `Show-Dashboard`  
  (after window initialization)
- StatusBar now empty but functional (no CLS compliance errors)
- All function keys still work through menu system

---

### 1.0.3.03  (Window Title Width Error Fix)
- **CRITICAL FIX:** Removed dynamic window title update  
  (causing *"Width must be >= 0"* error)
- Window title built from variables as always
- Fixes `Application::Run()` null toplevel error
- All DNS functionality works correctly
- CLS compliance errors eliminated
- Null array errors eliminated
