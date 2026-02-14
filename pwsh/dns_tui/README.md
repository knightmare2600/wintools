# easyDNS TUI

Terminal-based DNS management console for PowerShell using Terminal.Gui.  
Part of the **wintools PowerShell TUI suite**.

---

## Quick Start (30 seconds)

Run from PowerShell 7:

```powershell
pwsh ./dns_tui.ps1
```

If Terminal.Gui is missing:

```powershell
Install-Module Microsoft.PowerShell.ConsoleGuiTools
```

Ensure DNS tools are installed and you have permissions to query the DNS server.

---

## Table of Contents

- Overview
- Why This Tool Exists
- Core Features
- Requirements
- Installation Notes
- Running the Application
- Keyboard Navigation
- Common Problems & Fixes
- Security Notes
- Architecture Notes
- Intended Audience
- License
- Credits

---

## Overview

easyDNS TUI is a keyboard-driven DNS administration interface designed for
managing Microsoft DNS servers entirely from a terminal session.

It is intended primarily for:

- Windows Server Core systems
- Remote PowerShell / SSH sessions
- Environments without GUI tools
- Administrators who prefer console workflows

The application provides a structured text UI for browsing DNS zones,
managing records, performing diagnostics, and exporting DNS information.

---

## Why This Tool Exists

Microsoft DNS Manager (MMC):

- requires GUI access
- performs poorly over remote sessions
- is unavailable on Server Core
- is mouse-driven

easyDNS TUI provides:

- full keyboard control
- console-only operation
- compatibility with remote shells
- significantly lower overhead

---

## Core Features

### DNS Dashboard

- Connected DNS server display
- Zone statistics overview
- System information panel

---

### Zone Management

- Browse forward lookup zones
- Browse reverse lookup zones
- Create forward zones
- Create reverse zones
- Reload zones
- Trigger secondary transfers
- Refresh zone listings

---

### Record Management

- List records inside zones
- Create DNS records
- Delete DNS records
- Refresh record views

---

### DNS Diagnostics Toolkit

Built-in troubleshooting tools include:

- Ping tests
- Nslookup queries
- Resolve-DnsName lookups
- Traceroute functionality
- DNS performance benchmarking
- DNS cache inspection and clearing

---

### Import / Export

DNS information can be exported to:

- CSV
- JSON
- XML

Includes an integrated file-selection dialog.

---

## Requirements

### Required

- Windows system with DNS tools available
- PowerShell **7.x strongly recommended**
- DNS administrative permissions

### Terminal UI Framework

Requires Terminal.Gui assembly.

Install with:

```powershell
Install-Module Microsoft.PowerShell.ConsoleGuiTools
```

---

## ⚠ PowerShell Version Warning

Windows PowerShell 5.1 is **not recommended**.

Use:

```powershell
pwsh
```

PowerShell 7 provides:

- better runspace handling
- improved console rendering
- proper Unicode support
- more reliable Terminal.Gui behaviour

---

## Running the Application

Basic launch:

```powershell
pwsh ./dns_tui.ps1
```

If your configuration supports specifying a DNS server:

```powershell
pwsh ./dns_tui.ps1 -Server DC01
```

---

## Keyboard Navigation

| Key | Action |
|-----|--------|
| F1 | Help |
| F5 | Refresh |
| ESC | Close dialog |
| Ctrl+Q | Quit |

Menus provide additional navigation.

---

## Common Problems & Fixes

### Terminal.Gui Not Loading

Install ConsoleGuiTools:

```powershell
Install-Module Microsoft.PowerShell.ConsoleGuiTools
```

Restart PowerShell afterwards.

---

### DNS Cmdlets Missing

Install DNS tools via:

- Windows Server Roles → DNS Tools
- RSAT DNS Tools (client machines)

Verify:

```powershell
Get-Command Get-DnsServerZone
```

---

### Permission Errors

Run PowerShell elevated:

```
Run as Administrator
```

Ensure account is:

- DNS Admin
- Domain Admin
- or delegated DNS manager

---

### Rendering Issues Over Remote Sessions

If UI behaves oddly:

- resize terminal window
- reconnect session
- ensure UTF-8 console encoding
- avoid legacy Windows console host where possible

Windows Terminal is recommended.

---

## Security Notes

This tool operates against **live DNS infrastructure**.

- zone changes apply immediately
- record deletion is permanent
- always validate target server before modifying zones
- test in staging where possible

---

## Architecture Notes

The application is built using:

- PowerShell runspaces
- Terminal.Gui window system
- Modal dialog architecture
- Menu-driven command routing

Special defensive handling exists for:

- Terminal.Gui key enum issues
- status bar compatibility problems
- window sizing calculation bugs
- runspace timer conflicts

---

## Intended Audience

This tool is designed for:

- Windows infrastructure administrators
- enterprise sysadmins
- datacenter operators
- PowerShell-heavy operational teams
- Server Core environments

---

## Related wintools Components

Other TUI tools in the wintools suite include:

- DSA-TUI (Active Directory administration console)

---

## License

Released under **GPL-3.0**.

---

## Credits

- Terminal.Gui framework contributors  
- Microsoft DNS PowerShell tooling  
- Original easyDNS GUI inspiration
