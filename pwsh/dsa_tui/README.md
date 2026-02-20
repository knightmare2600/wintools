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

Some Screenshots of the script in action. For example, user properties modal with the DSB (Danske Statsbaner) theme:

![dsa_tui_001](screenshots/dsa_tui_001.png?raw=true "Example User Properties")

Some Easter eggs hidden in the data, here using the procomm theme:

![dsA_002](screenshots/dsa_tui_002.png?raw=true "Example Easter Egg")

It is also possible to query the properties of other AD objects, along with the Pan Am theme:

![dsa_003](screenshots/dsa_tui_003.png?raw=true "Example OU Properties")

The LAPS username and password viewer is included, to avoid having to remember esoteric powershell commands:

![dsA_004](screenshots/dsa_tui_004.png?raw=true "LAPS Viewer")

A complimentary password generator, this time in the gemstones theme:

![dsa_005](screenshots/dsa_tui_005.png?raw=true "Password Generator")

And lastly, we have device properties for an Internet connected Jukebox in Odense, Danmark with the 1980s DB Orient express theme:

![dsa_006](screenshots/dsa_tui_006.png?raw=true "Jukebox device properties")

The demo data for AD Forest `jukebox.example` is all stored inside the code, so can be executed safely.

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
| Jan 1  | 📅 | New Year's day. Start of a new calendar year | |
| Jan 2  | 🦄 | Tradiitonal day to hunt Wild Haggis in Scotland | https://haggiswildlifefoundation.com |
| Apr 9  | 🇩🇰 | German invasion of Denmark (1940) – start of occupation | https://en.wikipedia.org/wiki/German_invasion_of_Denmark_(1940) |
| May 4  | 🕯️ | Danish remembrance candle for the occupation (Besættelsen) | https://en.wikipedia.org/wiki/Denmark_in_World_War_II |
| Jun 5  | 🇩🇰 | Liberation of Denmark (1945) | https://en.wikipedia.org/wiki/Liberation_of_Denmark |
| Jun 21 | 🇬🇱 | Greenland National Day | https://en.wikipedia.org/wiki/Greenland_National_Day |
| Jul 1  | 🇨🇦 | Canada Day eh / Fête du Canada le eh  | https://en.wikipedia.org/wiki/Tea_in_the_United_Kingdom |
| Jul 29 | 🇫🇴 | Ólavsøka – Faroe Islands national festival | https://en.wikipedia.org/wiki/%C3%93lavs%C3%B8ka |
| Nov 9  | 🇩🇪 | Historical in-joke referencing Erich Honecker | https://en.wikipedia.org/wiki/Erich_Honecker |
| Nov 24 | 👑 | Historical in-joke referencing Prince Knud Of Denmark | https://en.wikipedia.org/wiki/Knud,_Hereditary_Prince_of_Denmark |
| Nov 30 | 🏴󠁧󠁢󠁳󠁣󠁴󠁿 | St Andrew’s Day | https://en.wikipedia.org/wiki/Saint_Andrew's_Day |
| Dec 24–25 | 🎄 | Christmas / Jul | https://en.wikipedia.org/wiki/Christmas |

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
| Example Music (Australia) Pty Ltd     | Australia            | Pty Ltd.              | Proprietary Limited Company |
| Example Music (Canada) Inc.           | Canada               | Inc.                  | Federal / provincial corporations |
| Example Creative ApS                  | Denmark              | ApS                   | Danmark Aktieselskab (public limited) |
| Example Music (Danmark) ApS           | Denmark              | ApS                   | Danmark Aktieselskab (public limited) |
| Example Music (Deutschland) GmbH      | Germany              | GmbH                  | Gesellschaft mit beschränkter Haftung |
| Example Music (England) Ltd           | England & Wales      | Ltd                   | Companies House jurisdiction |
| Example Music (Italia) S.p.a.         | Italy                | S.p.a.                | Società per azioni |
| Example Music (Nederland) B.V.        | Netherlands          | B.V.                  | bv (besloten vennootschap) |
| Example Music (New Zealand) Tāpui     | New Zealand          | Tāpui                 | Standard NZ corporation suffix |
| Example Music (Norge) ASA             | Norway               | ASA                   | Allmennaksjeselskap |
| Example Music (Österreich) GmbH       | Austria              | GmbH                  | Gesellschaft mit beschränkter Haftung |
| Example Music (Scotland) Ltd          | Scotland             | Ltd (SCxxxxxx)        | Scottish company numbers start with “SC” |
| Example Music (Sverige) AB            | Sweden               | AB                    | Aktiebolag |
| Example Music (US) LLC.               | USA                  | LLC.                  | Standard US corporation suffix |

## Site Names And subnets

**Demo Site Names and subnets**

| **Site** | **City / Location**          | **Country** | **Subnet**            | **Landline Range**                    | **Mobile Range** |
|---------:|------------------------------|-------------|-----------------------|---------------------------------------|------------------|
| ABR | Aberdeen, Scotland             | UK | 192.168.224.0/24 | +44 1224 496 0xxx               | +44 7700 900 2xxx |
| AKL | Auckland, New Zealand          | NZ | 192.168.93.0/24  | +64 9 300 0xxx                  | +64 21 900 2xxx |
| BIR | Birmingham, England            | UK | 192.168.121.0/24 | +44 121  496 0xxx               | +44 7700 900 2xxx |
| BON | Bonn, West Germany (FRG)       | DE | 192.168.228.0/24 | +49 228  555  xxx               | +49 211  xxx xxxx |
| BRD | West Berlin (FRG)              | DE | 192.168.113.0/24 | +49 311  555  xxx               | +49 211  xxx xxxx |
| BRK | Brockville, Ontario            | CA | 192.168.136.0/24 | +1  613  555 6xxx               | +1  613  555 6xxx |
| CLY | Clydebank, Scotland            | UK | 192.168.41.0/24  | +44 141  496 00xx               | +44 770  090 5xxx |
| COV  | Coventry, England             | UK | 192.168.247.0/24 | +44 247 765 0xxx                | +44 7700 901 2xxx |
| CPH | Copenhagen (København)         | DK | 192.168.231.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| DUN | Dundee, Scotland               | UK | 192.168.138.0/24 | +44 163  249 60xx               | +44 770  090 82xx |
| EDI | Edinburgh, Scotland            | UK | 192.168.131.0/24 | +44 131  496 0xxx               | +44 770  090 3xxx |
| FAL  | Falkirk, Scotland             | UK | 192.168.76.0/24  | +44 1324 500 0xxx               | +44 7700 903 2xxx |
| FAX | Faxe, Danmark                  | DK | 192.168.246.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| GAA  | Georgia, AL, USA              | US | 192.168.33.0/24  | +1 334 300 0xxx                 | +1 770 900 2xxx  |
| GLA | Glasgow, Scotland              | UK | 192.168.141.0/24 | +44 141  496 01xx               | +44 770  009 4xxx |
| HAL  | Halifax, England              | UK | 192.168.142.0/24 | +44 1422 200 0xxx               | +44 7700 904 2xxx |
| HUL  | Hull, England                 | UK | 192.168.148/24   | +44 1482 300 0xxx               | +44 7700 902 2xxx |
| KGE | Køge, Danmark                  | DK | 192.168.265.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| KOR | Korsør, Danmark                | DK | 192.168.238.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| LIV | Liverpool, England             | UK | 192.168.151.0/24 | +44 151  496 0xxx               | +44 770  090 5xxx |
| LAX | Los angeles, California        | US | 192.168.213.0/24 | +1 213  555 xxxx               | +1 213 555 xxx |
| LND | London, England                | UK | 192.168.20.0/24  | +44 207  496 0xxx / 01632 96x xxx | +44 770  090 0xxx |
| MCR | Manchester, England            | UK | 192.168.161.0/24 | +44 161  715 xxxx               | +44 770  090 6xxx |
| MEL | Melbourne, Australia           | AU | 192.168.39.0/24  | +61 3 9000 0xxx                 | +61 400 901 2xxx |
| MIA | Miami, Florida                 | US | 192.168.135.0/24 | +1  305  555 xxxx               | +1  786  555 xxxx |
| MTL | Montreal, Canada               | CA | 192.168.154.0/24 | +1 514 400 0xxx                 | +1 514 900 2xxx  |
| MUN | Munich, West Germany           | DE | 192.168.189.0/24 | +49 893  555 33xx               | +49 893  555 99xx |
| NEW | Newcastle, England             | UK | 192.168.191.0/24 | +44 191  496 0xxx               | +44 770  090 9xxx |
| NYC | New York, NY, USA              | US | 192.168.212.0/24 | +1 212 500 0xxx                 | +1 917 900 2xxx  |
| NJC | New Jersey (City), NJ, USA     | US | 192.168.201.0/24 | +1 201 400 0xxx                 | +1 908 900 2xxx  |
| ODE | Odense, Danmark                | DK | 192.168.126.0/24 | +45  00  000  xxx               | +45  2x  xxx xxx |
| PER | Perth, Scotland                | UK | 192.168.173.0/24 | +44 173  849 60xx               | +44 770  0173 0xx |
| SHE | Sheffield, England             | UK | 192.168.114.0/24 | +44 114 250 0xxx                | +44 7700 905 2xxx |
| SYD | Sydney, Australia              | AU | 192.168.29.0/24  | +61 2 9000 0xxx                 | +61 400 900 2xxx |
| VIE | Vienna, Austria                | AT | 192.168.78.0/24  | +43 800  078  0xx               | +43 664  665 xxx |

TODO: Oslo, Gothemburg, San Francisco, Chicago, Amsterdam, Seattle, 

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
**__NB: I don't believe Denmark uses 0000 but this is not confirmed!__**

## Other information relating to the demo data:

Yes, Christianshavn alludes to [this](https://en.wikipedia.org/wiki/Huset_p%C3%A5_Christianshavn?useskin=vector) Danish TV show and
[Helena Christensen](https://en.wikipedia.org/wiki/Helena_Christensen?useskin=vector) is the lady from the Chris Issak music video and you'll find other such Easter Eggs in the code and AD Data.

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
MAC  = macOS Desktop (iMac / Mac Mini)                        TAR  = Tape archiver (Yes, it's where UNIX tar comes from
MBP  = MacBook Pro                                            TEA  = Internet connected Coffee Pot / Tea machine RFC2324 compliant
MUS  = Music Workstation / Studio System / Jukebox            TTY  = Teletype termianl E.g DEC VT320
NIX  = Unix/Linux/Solaris System                              TVS  = Television / Digital Signage Display
PBX  = PBX (Telephone server)                                 VCU  = Video Conferencing Unit
PHN  = Mobile Phone                                           VND  = Vendinh Machine
PMP  = Petrol Pump                                            WAP  = Wireless Access Point
PRN  = Printer                                                WKS  = Workstation (Desktop)
```

Device Numbering: 000–999 (three digits, zero padded)

Examples:
```
  EXAWKSBON001  -> Bonn workstation
  EXASRVEDI003  -> Edinburgh server
  EXAWAPLND001  -> London Wi-Fi access point
```

###

Railway Maps giving an idea of office locations in relation to each other:

Danmark: DSB network map:

![DK_map](screenshots/DSB_Map.png?raw=true "The day we caught the train")

Germany:

![DE_map](screenshots/DB1980s.jpg?raw=true "Trans Europe express")

Canada eh:

![CA_map](screenshots/Viarail80s.png?raw=true "Take off to the Great White North")

Scotland:

![Scotland_map](screenshots/ScotFail.jpeg?raw=true "I saw you up on a clear day. First taking hearts. Then our last breath away")

England:

![BR_map](screenshots/BREngland.jpg?raw=true "Found myself in a strange town")

This may help with the Geography.

## Additional Notes

## LAPS

## 📘 LAPS Handling & Application Flow — Technical Documentation

---

## 🔐 LAPS Versions Explained

### Legacy LAPS (2015)
Legacy Microsoft LAPS stores credentials using the following Active Directory attributes:

- **`ms-Mcs-AdmPwd`**  
  Stores the local administrator password (plaintext in AD, ACL-protected).

- **`ms-Mcs-AdmPwdExpirationTime`**  
  Stores password expiry as a Windows **FileTime** value.


### Windows LAPS (2023+)
Modern Windows LAPS introduces a new schema with clearer semantics and better extensibility:

- **`msLAPS-Password`**  
  Stores the managed local administrator password.

- **`msLAPS-PasswordExpirationTime`**  
  Password expiration stored as **FileTimeUtc**.

- **`msLAPS-AccountName`**  
  Name of the managed local account (defaults to `Administrator`).

---

## 🧩 TDF Property Mappings

The **TDF (TUI DSA Format)** file uses Windows LAPS–style properties internally.

### Raw Properties
- `msLAPS-Password` → `Bon#Lap@64`
- `msLAPS-PasswordExpirationTime` → `FileTimeUtc`
- `msLAPS-AccountName` → `Administrator`


## 🔍 LAPS Detection Order

The application resolves LAPS data using a strict priority order to ensure compatibility across environments:

1. **Windows LAPS attributes**  
   (`msLAPS-*`)

2. **Friendly aliases**  
   (`LAPSPassword`, `LAPSPasswordExpiration`, etc.)

3. **Legacy LAPS attributes**  
   (`ms-Mcs-*`)

This guarantees:
- Forward compatibility
- Backward compatibility
- Seamless demo and production behaviour

---

### Friendly Aliases

To simplify UI and scripting, the following aliases are also supported:

- **`LAPSPassword`**			Alias for `msLAPS-Password`
- **`LAPSPasswordExpiration`**  	Expiration converted to `DateTime`
- **`LAPSPasswordLastSet`**		Derived `DateTime` value for display and auditing

## 🧠 Summary of Code Changes

### ✅ What Changed
The **only** changes compared to the original implementation are:

1. **Progress percentages** added to `Set-StatusBar` calls  
2. **Verification log** added immediately after data load  
3. **Object count log** added before tree construction  

### ❌ What Did *Not* Change
Everything else remains **exactly the same**, including:

- Menu system
- Filter panel
- Selection panel
- Information panel
- Tree structure
- Key handlers
- UI layout
- Navigation logic

No refactors. No behavioural drift. No regressions.

## 🔁 Application Startup Order

The application always executes in the following order:

1. Load data (**Step 6**)
2. Create menu
3. Create filter panel
4. Create selection panel
5. Create info panel
6. Build tree
7. Add key handlers
8. Run application

This order is **intentional and fixed** to preserve predictable UI behaviour and startup performance.

Notes:
  - Not all devices are AD-aware
  - Some non-AD devices may still have service accounts
  - LAPS attributes exist only on supported Windows computers
  - Some parts of the demo AD data are intentionally broken and can be "fixed" if you are learning
