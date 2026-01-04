<!-- TOC -->
- [About This Project](#about-this-project)
- [Requirements](#requirements)
- [Parameters](#parameters)
  - [CSV Import Notes](#csv-import-notes)
- [Required PowerShell Modules](#required-powershell-modules)
  - [ActiveDirectory (RSAT)](#activedirectory-rsat)
- [Citations for the Active Directory Commands](#citations-for-the-active-directory-commands)
- [Example Common LDAP Filters (for reference/docs)](#example-common-ldap-filters-for-referencedocs)
- [Demo Data And Other Relevant Information](#demo-data-and-other-relevant-information)
  - [Special Days & Emoji Easter Eggs](#special-days--emoji-easter-eggs)
- [Notes](#notes)
- [Information One May Find Useful](#information-one-may-find-useful)
  - [Companies](##Companies)
  - [Demo Data And Other Relevant Information](#Demo-Data-And-Other-Relevant-Information)
  - [Special Days And easter Eggs](#Special-Days-And-Emoji-Easter-Eggs)
  - [Device Naming & Demo Data Conventions](#device-naming--demo-data-conventions)
  - [Company Suffixes](#Companies)
  - [Demo Site Names and Subnets](#demo-site-names-and-subnets)
  - [Device Role Codes](#Device-Naming-Conventions)
  - [Notes](#Additional-Notes)
<!-- /TOC -->

# About This project

DSA-TUI Text Mode version of dsa.msc for powershell using ConsoleTools 1.16 on Powershell 7.x

# Requirements

- Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup
- Powershell 7+
- ConsoleTools Powershell module is __MANDATORY__ others, including ActiveDirectory are useful, but optional

__If you don't have them, you can run in Demo mode__

# dsa_tui.ps1 — Parameters & Usage

The `dsa_tui.ps1` script supports several command-line parameters to control behaviour, data sources, and targeting of Active Directory environments.

---

## Parameters

| **Command Line Parameter** | **Type** | **Description** | **Example Usage In Script** |
|---------------|----------|------------------|-------------|
| `-ImportCsv`  | `string` | Path to a CSV file containing objects to import into Active Directory. The CSV should follow a standard `csvde` layout (attribute names as headers). | `.\dsa_tui.ps1 -ImportCsv C:\temp\users.csv` |
| `-Domain` | `string` | Specifies the Active Directory domain to operate against. If omitted, the current logon domain is used. | `.\dsa_tui.ps1 -Domain corp.example.com` |
| `-DC` | `string` | Explicitly specifies a domain controller. Useful in multi-site or troubleshooting scenarios. | `.\dsa_tui.ps1 -DC dc01.corp.example.com` |
| `-Verbose` | `switch` | Enables verbose output during execution. | `.\dsa_tui.ps1 -Verbose` |
| `-Help` | `switch` | Displays usage information. | `.\dsa_tui.ps1 -Help` |

---

## CSV Import Notes

When using `-ImportCsv`, the CSV must already exist. The script does **not** generate or infer CSV structure.

A common workflow is to export existing objects using `csvde`:

```powershell
csvde -f exported_users.csv
```

More information on the `csvde.exe` tool can be found [here](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/cc732101(v=ws.11))

---

## Required PowerShell Modules

### ActiveDirectory (RSAT)

Install RSAT:

    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-Lite~~~~0.0.1.0

Load the module:

    Import-Module ActiveDirectory

---

## Citations for the Active directory commands:

  Reference URLs for AD Forest & Get-ADForest usage:

  - Microsoft Official Documentation
    [Get-ADForest](https://learn.microsoft.com/powershell/module/activedirectory/get-adforest)
    [AD Forest and domain guide](https://learn.microsoft.com/windows-server/identity/ad-ds/plan/active-directory-forest-and-domain-guide)

  - General AD Forest Structure Explanations
    [Microsoft: Undderstanding AD](https://learn.microsoft.com/windows-server/identity/ad-ds/plan/understanding-active-directory-domain-services)
    [Microsoft: Designing AD Structure] (https://learn.microsoft.com/windows-server/identity/ad-ds/plan/designing-the-logical-structure)

  - Community / Confirmatory Sources
    [SS64 AD](https://ss64.com/ps/get-adforest.html)
    [4sysops](https://4sysops.com/archives/get-adforest-cmdlet-what-it-does-and-how-to-use-it/)
    [Technet AD forest discussion](https://social.technet.microsoft.com/wiki/contents/articles/22953.active-directory-forest.aspx)

Notes:
```
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
```

## Example Common LDAP Filters (for reference/docs)

Common LDAP Filters:

```
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
```

# Demo Data And Other Relevant Information

Here you will find some infromation relating to the dmeo data and quesitons you may have regarding it. Along with some fun Easter eggs

## Special Days And Emoji Easter Eggs

| **Date** | **Emoji** | **Description** | **Reference** |
|---------:|-----------|-----------------|---------------|
| April 9 | 🇩🇰 | German invasion of Denmark (1940) – start of occupation | https://en.wikipedia.org/wiki/German_invasion_of_Denmark_(1940) |
| May 4 | 🕯️ | Danish remembrance candle for the occupation (Besættelsen) | https://en.wikipedia.org/wiki/Denmark_in_World_War_II |
| June 5 | 🇩🇰 | Liberation of Denmark (1945) | https://en.wikipedia.org/wiki/Liberation_of_Denmark |
| June 21 | 🇬🇱 | Greenland National Day | https://en.wikipedia.org/wiki/Greenland_National_Day |
| July 1 | 🇨🇦 | Canada Day eh / Fête du Canada le eh  | https://en.wikipedia.org/wiki/%C3%93lavs%C3%B8ka](https://en.wikipedia.org/wiki/Canada_Day |
| July 4 | 🫖 | Drink a cup of tea | https://en.wikipedia.org/wiki/%C3%93lavs%C3%B8ka](https://en.wikipedia.org/wiki/Tea_in_the_United_Kingdom |
| July 29 | 🇫🇴 | Ólavsøka – Faroe Islands national festival | https://en.wikipedia.org/wiki/%C3%93lavs%C3%B8ka |
| November 9 | 🇩🇪 | Historical in-joke referencing Erich Honecker | https://en.wikipedia.org/wiki/Erich_Honecker |
| November 24 | 👑 | Historical in-joke referencing Prince Knud Of Denmark | https://en.wikipedia.org/wiki/Knud,_Hereditary_Prince_of_Denmark |
| November 30 | 🏴󠁧󠁢󠁳󠁣󠁴󠁿 | St Andrew’s Day | https://en.wikipedia.org/wiki/Saint_Andrew's_Day |
| December 24–25 | 🎄 | Christmas / Jul | https://en.wikipedia.org/wiki/Christmas |

Default emoji on all other days: 🗂️

For November 9th, reference, check out the movie _The Lives of Others_ — [here](https://www.imdb.com/title/tt0405094/)

---

## Notes

- CSV files must already exist before import.
- Attribute names must match the Active Directory schema.
- RSAT / ActiveDirectory module is required.
- Emoji easter eggs are cosmetic only.

## Information one may find useful:

Some information which might come in handy: [Someone else who likes the projct name](https://nyheder.tv2.dk/lokalt/2021-10-19-er-det-her-postbuddets-vaerste-skraek-hvem-pokker-har-fundet-paa-det-her)

## Companies

<ins>DEVICE NAMING & DEMO DATA CONVENTIONS</ins>

**Company Suffixes:**

| **Name**                              | **Country**          | **Legal Form**        | **Notes** |
|--------------------------------------|----------------------|-----------------------|-----------|
| Example Music (England) Ltd           | England & Wales      | Ltd                   | Companies House jurisdiction |
| Example Music (Scotland) Ltd          | Scotland             | Ltd (SCxxxxxx)        | Scottish company numbers start with “SC” |
| Example Music (Danmark) ApS           | Denmark              | ApS                   | Danmark Aktieselskab (public limited) |
| Example Music (Deutschland) GmbH      | Germany              | GmbH                  | Gesellschaft mit beschränkter Haftung |
| Example Music (Österreich) GmbH       | Austria              | GmbH                  | Gesellschaft mit beschränkter Haftung |
| Example Music (CA) Inc.               | Canada               | Inc.                  | Federal / provincial corporations |
| Example Music (US) LLC.               | USA                  | LLC.                  | Standard US corporation suffix |

## Site Names And subnets

**Demo Site Names and subnets**

| **Site** | **City / Location**          | **Country** | **Subnet**            | **Landline Range**                    | **Mobile Range** |
|---------:|------------------------------|-------------|-----------------------|---------------------------------------|------------------|
| ABR | Aberdeen, Scotland             | UK | 192.168.224.0/24 | +44 1224 496 0xxx               | +44 7700 900 2xxx |
| BIR | Birmingham, England            | UK | 192.168.121.0/24 | +44 121  496 0xxx               | +44 7700 900 2xxx |
| BON | Bonn, West Germany (FRG)       | DE | 192.168.228.0/24 | +49 228  555  xxx               | +49 211  xxx xxxx |
| BRD | West Berlin (FRG)              | DE | 192.168.113.0/24 | +49 311  555  xxx               | +49 211  xxx xxxx |
| BRK | Brockville, Ontario            | CA | 192.168.136.0/24 | +1  613  555 6xxx               | +1  613  555 6xxx |
| CLY | Clydebank, Scotland            | UK | 192.168.141.0/24 | +44 141  496 00xx               | +44 770  090 5xxx |
| CPH | Copenhagen (København)         | DK | 192.168.231.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| DUN | Dundee, Scotland               | UK | 192.168.138.0/24 | +44 163  249 60xx               | +44 770  090 82xx |
| EDI | Edinburgh, Scotland            | UK | 192.168.131.0/24 | +44 131  496 0xxx               | +44 770  090 3xxx |
| FAX | Faxe, Danmark                  | DK | 192.168.246.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| GLA | Glasgow, Scotland              | UK | 192.168.141.0/24 | +44 141  496 01xx               | +44 770  009 4xxx |
| KGE | Køge, Danmark                  | DK | 192.168.265.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| KOR | Korsør, Danmark                | DK | 192.168.238.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| LIV | Liverpool, England             | UK | 192.168.151.0/24 | +44 151  496 0xxx               | +44 770  090 5xxx |
| LND | London, England                | UK | 192.168.20.0/24  | +44 207  496 0xxx / 01632 96x xxx | +44 770  090 0xxx |
| MCR | Manchester, England            | UK | 192.168.161.0/24 | +44 161  715 xxxx               | +44 770  090 6xxx |
| MIA | Miami, Florida                 | US | 192.168.135.0/24 | +1  305  555 xxxx               | +1  786  555 xxxx |
| MUN | Munich, West Germany           | DE | 192.168.189.0/24 | +49 893  555 33xx               | +49 893  555 99xx |
| NEW | Newcastle, England             | UK | 192.168.191.0/24 | +44 191  496 0xxx               | +44 770  090 9xxx |
| ODE | Odense, Danmark                | DK | 192.168.126.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| PER | Perth, Scotland                | UK | 192.168.173.0/24 | +44 173  849 60xx               | +44 770  0173 0xx |
| VIE | Vienna, Austria                | AT | 192.168.xxx.0/24 | +43 800  078  0xx               | +43 664  665 xxx |

**__A note on phone numbers:__**

```
  - UK Phone Number Standards (Ofcom reserved ranges for fiction/testing):
    - Glasgow:                 0141 496 0xxx
    - Edinburgh:               0131 496 0xxx
    - London:                  0207 946 0xxx / 01632 96x xxx
    - Manchester:              0161 715 xxxx - Weatherfield
    - UK Wide Mobiles: Mobile: 0770 090 0xxx
```
  Per: [Ofcom Fictitious_numbers](https://en.wikipedia.org/wiki/Telephone_numbers_in_the_United_Kingdom?useskin=vector#Fictitious_numbers)
  and: [Numbers For Drama](https://web.archive.org/web/20140216214400/http://stakeholders.ofcom.org.uk/telecoms/numbering/guidance-tele-no/numbers-for-drama)

**Yes, 0161 715 is a call back to [Coronation Street](https://en.wikipedia.org/wiki/Coronation_Street?useskin=vector) per the ["Coronation Street test" procedure](https://en.wikipedia.org/wiki/TV_pickup?useskin=vector)**

```
  - Denmark Testing Numbers:
    - Copenhagen landline: +45 0000 xxxx
    - Denmark mobile:      +45 2xxx xxxx
```
## Other information relating to the demo data:

Yes, Christianshavn alludes to [this](https://en.wikipedia.org/wiki/Huset_p%C3%A5_Christianshavn?useskin=vector) Danish TV show and
[Helena Christensen](https://en.wikipedia.org/wiki/Helena_Christensen?useskin=vector) is the lady from the Chris Issak music video and you'll find other such Easter Eggs in the code and AD Data.

**__NB: I don't believe Denmark uses 0000 but this is not confirmed!__**

## Device Naming Conventions

Device Role Codes:
```
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
```

Device Numbering: 000–999 (three digits, zero padded)

Examples:
```
  EXAWKSBON001  -> Bonn workstation
  EXASRVEDI003  -> Edinburgh server
  EXAWAPLND001  -> London Wi-Fi access point
```

## Additional Notes

Notes:
  - Not all devices are AD-aware
  - Some non-AD devices may still have service accounts
  - LAPS attributes exist only on supported Windows computers
  - Some parts of the demo AD data are intentionally broken and can be "fixed" if you are learning
