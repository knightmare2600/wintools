<# History

v1.0.0 – Show-FileBrowser Modal initial version
- Fixed dark mode, modal behavior, and inline selection (compatible with Terminal.Gui v1.16+).

v1.1.0 – Allow PEM files to be read
- Added ability for PEM fiels to be read
- Replaced Application.Run calls with RunIteration loops for cross-version Terminal.Gui stability.

v1.2.0 – Validations
- Added mandatory email domain check against CN. Not just if the address is valid.
- Implemented standard email format verification.
- cosmetic fixes on New Cert dialog including sanity checks for country codes and email address
- Added inline [?] help button for email validation.
- Corrected vertical alignment of email validation button.
- Preserved original layout while shortening email field minimally.

v1.4.0 – Wildcard & Domain Handling
- Properly handles wildcard CNs.
- Warns users if email domain does not match CN.
- Catch common mistakes such as picking UK instead of GB for country code

v1.4.1 – Wildcard & Domain Handling
- F4 shortcut for cert info
- Fix regression with wildcard handling
- Keep dialogs on screen after generating files, so you can have OpenSSL and INF if you so desire.

v1.5.8
- Add certificate conversion modal giving openssl, certutil and powershell helpers
- Modal will, if it finds openSSL in the path, offer to run the commands
- Password support for p7bs and when generating
- File browser debug logging
- Private key support
- Download Google's SSL with certutil or openssl for user test conversions
- Fix logic bug inside file browser so path "falls through" to script
- Make themes more defined in their colour schemes
- Modals are now clickable and copyable for things such as URLs
- Easter egg hidden faxe kondi theme mode - Matches Faxe Kondi livery of green, yellow, and grey

TODO
- ability to open a CSR file and parse it for valueas to furnish a new cnf/inf (complex)
- given a certificate chain, offer to split into intermediates and ssl (some software such as Acronis needs this)

#Requires - Version 7.0+
#>

<#
.SYNOPSIS
    PowerShell SSL Certificate Helper with Terminal.Gui
.NOTES
    Version: ${BuildVersion} STABLE
    Terminal.Gui: v1.16.0
    
    Certificate generation code adapted from:
    Author: Roberto Rodriguez (@Cyb3rWard0g)
    Source: https://raw.githubusercontent.com/OTRF/Blacksmith/master/resources/scripts/powershell/misc/Get-CertSigningReq.ps1
    License: GPL-3.0
#>

param(
    [switch]$Verbose,
    [ValidateSet("Dark","Light","Faxekondi","British","Default")]
    [string]$Theme = "Dark"
)

# Define the build version once
$BuildVersion = "1.5.8"

function Debug-Log {
    param([string]$Message)
    if ($Verbose) {
        $ts = (Get-Date).ToString("HH:mm:ss")
        Write-Host "[$ts] LOG: $Message" -ForegroundColor Cyan
    }
}

# Load Terminal.Gui
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq "Terminal.Gui" })) {
    $mod = Get-Module Microsoft.PowerShell.ConsoleGuiTools -List
    if (-not $mod) { throw "ConsoleGuiTools module not found. Install: Install-Module Microsoft.PowerShell.ConsoleGuiTools" }
    $dll = Join-Path $mod.ModuleBase "Terminal.Gui.dll"
    Add-Type -Path $dll
    
    if ($Verbose) {
        $asmVer = [System.Reflection.AssemblyName]::GetAssemblyName($dll).Version
        Debug-Log "Terminal.Gui version: $asmVer"
        Debug-Log "Module version: $($mod.Version)"
    }
}

[Terminal.Gui.Application]::Init()

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
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Green)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
        }

        "faxekondi" {
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::DarkGray)
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

## Confirm country code is valid. We don't use UK it's GB and so on....
function Get-ValidCountryCodes {
    return @(
        "AD", "ae", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AX", "AZ",
        "BA", "BB", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ", "BR", "BS", "BT", "BV", "BW", "BY", "BZ",
        "CA", "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN", "CO", "CR", "CU", "CV", "CW", "CX", "CY", "CZ",
        "DE", "DJ", "DK", "DM", "DO", "DZ",
        "EC", "EE", "EG", "EH", "ER", "ES", "ET",
        "FI", "FJ", "FK", "FM", "FO", "FR",
        "GA", "GB", "GD", "GE", "GF", "GG", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS", "GT", "GU", "GW", "GY",
        "HK", "HM", "HN", "HR", "HT", "HU",
        "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IR", "IS", "IT",
        "JE", "JM", "JO", "JP",
        "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ",
        "LA", "LB", "LC", "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY",
        "MA", "MC", "MD", "ME", "MF", "MG", "MH", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ",
        "NA", "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU", "NZ",
        "OM",
        "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PM", "PN", "PR", "PS", "PT", "PW", "PY",
        "QA",
        "RE", "RO", "RS", "RU", "RW",
        "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS", "ST", "SV", "SX", "SY", "SZ",
        "TC", "TD", "TF", "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TO", "TR", "TT", "TV", "TW", "TZ",
        "UA", "UG", "UM", "US", "UY", "UZ",
        "VA", "VC", "VE", "VG", "VI", "VN", "VU",
        "WF", "WS",
        "YE", "YT",
        "ZA", "ZM", "ZW"
    )
}

## Check the email matches the CN to avoid isses like WoSign https://blog.mozilla.org/security/2016/10/24/distrusting-new-wosign-and-startcom-certificates/
function Validate-EmailAgainstCN {
param(
[Parameter(Mandatory=$true)][string]$Email,
[Parameter(Mandatory=$true)][string]$CN
)
## Extract domain from email
if ($Email -notmatch "@(.+)$") { return $false }
  $emailDomain = $Matches[1].ToLower()

  # Normalize CN
  $cnLower = $CN.ToLower()

  # Handle wildcard CN (*.example.com)
  if ($cnLower.StartsWith("*.")) {
    $cnDomain = $cnLower.Substring(2)  # remove *.
    return $emailDomain.EndsWith($cnDomain)
}

## Exact match
return $emailDomain -eq $cnLower
}

## Show popup modals
function Show-Modal {
    param(
        [string]$Title,
        [string]$Message,
        [switch]$AllowCopyPaste,
        [switch]$CenterText
    )

    ## Split the message into lines
    $lines = $Message -split "`n"

    ## Compute dialog size
    $width  = ($lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum + 6
    $width  = [Math]::Min($width, [Terminal.Gui.Application]::Driver.Cols - 4)
    $height = $lines.Count + 6

    ## Create dialog
    $dialog = [Terminal.Gui.Dialog]::new($Title, $width, $height)

    if ($AllowCopyPaste) {
        ## Read-only TextView for copy/paste
        $textView = [Terminal.Gui.TextView]::new(2,1,$width-4,$height-4)
        $textView.ReadOnly = $true
        $textView.Text = [NStack.ustring]::Join("`n", $lines)
        $dialog.Add($textView)
    }
    else {
        ## Labels for normal modal
        $y = 1
        foreach ($line in $lines) {
            if ($CenterText) {
                $x = [Math]::Max(0, [Math]::Floor(($width - $line.Length)/2))
            } else {
                $x = 2
            }
            $dialog.Add([Terminal.Gui.Label]::new($x, $y, $line))
            $y++
        }
    }

    ## Center OK button
    $btnWidth = 8
    if ($CenterText) {
        # Vertically center button as well
        $btnX = [Math]::Floor(($width - $btnWidth)/2)
        $btnY = $height - 3
    } else {
        $btnX = 2
        $btnY = $height - 3
    }
    $okBtn = [Terminal.Gui.Button]::new($btnX, $btnY, "OK")
    $okBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dialog.Add($okBtn)

    # Run dialog
    [Terminal.Gui.Application]::Run($dialog)
}

## Show info about the project
function Show-JordbaerInfo {
  Show-Modal "Why Jordbaer...?" "certUI is codenamed Jordbaer because:`n`n- I was drinking Faxe Kondi Jordbaer when writing this code.`n- Jordbaer is Danish for Strawberry.`n- This is becoming a theme of fruit-based project names." -CenterText -EnableCopy
}

## Let users work with Google certificate if they wish to practice
function Show-DownloadExampleCertDialog {
##    $dialog = [Terminal.Gui.Dialog]::new("Download Example Certificate", 90, 50) ## <-- old code as a fallback
    $termHeight = [Console]::WindowHeight
    $dlgHeight = [Math]::Floor($termHeight * 0.85)  # 90% of terminal height
    $dialog = [Terminal.Gui.Dialog]::new("Download Example Certificate", 90, $dlgHeight)

    ## --- Tool availability checks ---
    $hasOpenSSL = $null -ne (Get-Command openssl -ErrorAction SilentlyContinue)
    $hasCertUtil = $null -ne (Get-Command certutil -ErrorAction SilentlyContinue)

    $y = 1

    ## --- OS + Tool selection side-by-side ---
    $labelOS = [Terminal.Gui.Label]::new(2, $y, "Select platform:")
    $dialog.Add($labelOS)

    $osOptions = @("Windows","Linux","macOS")
    $radioOS = [Terminal.Gui.RadioGroup]::new(20, $y, $osOptions)
    if ($IsWindows) { $radioOS.SelectedItem = 0 }
    elseif ($PSStyle.Platform -match "Linux") { $radioOS.SelectedItem = 1 }
    else { $radioOS.SelectedItem = 2 }
    $dialog.Add($radioOS)

    ## same row, right side
    $labelToolChoice = [Terminal.Gui.Label]::new(45, $y, "Select tool:")
    $dialog.Add($labelToolChoice)

    $toolOptions = @()
    if ($hasOpenSSL) { $toolOptions += "OpenSSL" }
    if ($IsWindows -and $hasCertUtil) { $toolOptions += "certutil" }
    $toolOptions += "PowerShell"
    $toolOptions += "All"

    $radioToolChoice = [Terminal.Gui.RadioGroup]::new(57, $y, $toolOptions)
    $radioToolChoice.SelectedItem = 0
    $dialog.Add($radioToolChoice)

    ## add one blank line visually
    $y += 4

    ## --- Available Tools Section ---
    $labelAvailable = [Terminal.Gui.Label]::new(2, $y, "Available tools:")
    $dialog.Add($labelAvailable)

    ## Build tool availability list
    $toolStatus = @()

    ## OpenSSL
    if ($hasOpenSSL) {
        $toolStatus += "OpenSSL: ✔  Available"
    } else {
        $toolStatus += "OpenSSL: ✖  Unvailable"
    }

    ## certutil (Windows only)
    if ($IsWindows) {
        if ($hasCertUtil) {
            $toolStatus += "certutil: ✔  Available"
        }
        else {
            $toolStatus += "certutil: ✖  Unvailable"
        }
    }

    ## PowerShell (always available)
    $toolStatus += "PowerShell: ✔  Available"

    ## Print tool status lines
    $toolStatus | ForEach-Object {
    $y += 1
    $dialog.Add([Terminal.Gui.Label]::new(4, $y, $_))
    }

    # Add clean spacing after this block
    $y += 2

    ## --- Download section ---
    $labelDownload = [Terminal.Gui.Label]::new(2, $y, "Download certificate from:")
    $dialog.Add($labelDownload)
    $y += 1

    $textHost = [Terminal.Gui.TextField]::new(2, $y, 40, "www.google.com")
    $dialog.Add($textHost)

    $btnDownload = [Terminal.Gui.Button]::new(45, $y, "Download Certificate")
    $btnDownload.add_Clicked({
        $hostname = $textHost.Text.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($hostname)) { Show-Modal "Error" "Please enter a hostname"; return }
        $pemFile = "$hostname.pem"
        $cerFile = "$hostname.cer"
        if ((Test-Path $pemFile) -or (Test-Path $cerFile)) {
            $result = [Terminal.Gui.MessageBox]::Query("File Exists","Certificate files already exist. Overwrite?",@("Yes","No"))
            if ($result -ne 0) { return }
        }
        try {
            $selectedOS = @("Windows","Linux","macOS")[$radioOS.SelectedItem]
            if ($selectedOS -eq "Windows" -and $hasCertUtil) {
                $url = "https://$hostname"
                $output = & certutil -urlcache -split -f $url $cerFile 2>&1
                if ($LASTEXITCODE -eq 0) { Show-Modal "Success" "Certificate downloaded as $cerFile (DER format)`n`nYou can now convert it using the converter below." }
                else { throw "certutil failed: $output" }
            } elseif ($hasOpenSSL) {
                $output = echo Q | openssl s_client -connect "${hostname}:443" -servername $hostname 2>$null | openssl x509 -outform PEM -out $pemFile
                if ($LASTEXITCODE -eq 0) { Show-Modal "Success" "Certificate downloaded as $pemFile (PEM format)`n`nYou can now convert it using the converter below." }
                else { throw "OpenSSL failed: $output" }
            } else {
                Show-Modal "Error" "No certificate download tools available (OpenSSL or certutil). You can still provide a local certificate file."
            }
        } catch { Show-Modal "Error" "Failed to download certificate:`n$($_.Exception.Message)" }
    })
    $dialog.Add($btnDownload)
    $y += 2

    ## --- Separator ---
    $labelSep = [Terminal.Gui.Label]::new(2, $y, "─" * 84)
    $dialog.Add($labelSep)
    $y += 1

    ## --- Converter section ---
    $labelConverter = [Terminal.Gui.Label]::new(2, $y, "Certificate Converter:")
    $dialog.Add($labelConverter)
    $y += 2

    ## --- Input file ---
    $labelInput = [Terminal.Gui.Label]::new(2, $y, "Input file:")
    $textInputFile = [Terminal.Gui.TextField]::new(15,$y,35,"")
    $btnBrowseInput = [Terminal.Gui.Button]::new(57,$y,"Browse...")
    $btnBrowseInput.add_Clicked({
        Write-Verbose "[DEBUG] Browse Input clicked"
        $selected = Show-FileBrowserDialog -StartDir "." -Title "Select Input Certificate" -Filter @("*.pem","*.crt","*.cer","*.der","*.pfx","*.p12","*.*")
        Write-Verbose "[DEBUG] Selected Input: $selected"
        if ($selected) { 
            $textInputFile.Text = [NStack.ustring]::Make($selected)
            Update-ConversionCommand
        }
    })
    $dialog.Add($labelInput)
    $dialog.Add($textInputFile)
    $dialog.Add($btnBrowseInput)
    $y += 2

    ## --- Output file ---
    $labelOutput = [Terminal.Gui.Label]::new(2,$y,"Output file:")
    $textOutputFile = [Terminal.Gui.TextField]::new(15,$y,35,"")
    $btnBrowseOutput = [Terminal.Gui.Button]::new(57,$y,"Browse...")
    $btnBrowseOutput.add_Clicked({
        Write-Host "[DEBUG] Browse Output clicked"
        $selected = Show-FileBrowserDialog -StartDir "." -Title "Select Output File" -Filter @("*.*")
        Write-Host "[DEBUG] Selected Output: $selected"
        if ($selected) { $textOutputFile.Text = [NStack.ustring]::Make($selected); Update-ConversionCommand }
    })
    $dialog.Add($labelOutput); $dialog.Add($textOutputFile); $dialog.Add($btnBrowseOutput)
    $y += 2

    ## --- Input + Output format (side-by-side) ---
    $labelInputFormat  = [Terminal.Gui.Label]::new(2,  $y, "Input:")
    $labelOutputFormat = [Terminal.Gui.Label]::new(40, $y, "Output:")
    $dialog.Add($labelInputFormat)
    $dialog.Add($labelOutputFormat)
    $y += 1

    $inputFormats  = @("PEM","DER","PKCS12")
    $outputFormats = @("PEM","DER","PKCS12")

    $radioInputFormat  = [Terminal.Gui.RadioGroup]::new(2,  $y, $inputFormats)
    $radioOutputFormat = [Terminal.Gui.RadioGroup]::new(40, $y, $outputFormats)

    $radioInputFormat.SelectedItem  = 0
    $radioOutputFormat.SelectedItem = 0

    $dialog.Add($radioInputFormat)
    $dialog.Add($radioOutputFormat)

    $y += 4   ## total vertical space to continue layout

    ## --- Key file & password ---
    $labelKeyFile = [Terminal.Gui.Label]::new(2,$y,"Private key (for PEM→PFX or DER→PFX):")
    $textKeyFile = [Terminal.Gui.TextField]::new(40,$y,25,"")
    $btnBrowseKey = [Terminal.Gui.Button]::new(68,$y,"Browse...")
    $btnBrowseKey.add_Clicked({
        $selected = Show-FileBrowserDialog -StartDir "." -Title "Select Private Key File" -Filter @("*.key","*.pem","*.*")
        if ($selected) { $textKeyFile.Text = [NStack.ustring]::Make($selected); Update-ConversionCommand }
    })
    $dialog.Add($labelKeyFile); $dialog.Add($textKeyFile); $dialog.Add($btnBrowseKey)
    $y += 2

    $labelPassword = [Terminal.Gui.Label]::new(2,$y,"Password (for PFX export):")
    $textPassword = [Terminal.Gui.TextField]::new(40,$y,25,"")
    $textPassword.Secret = $true
    $dialog.Add($labelPassword); $dialog.Add($textPassword)
    $y += 2

    ## --- Command preview ---
    $labelCommand = [Terminal.Gui.Label]::new(2,$y,"Commands preview:")
    $dialog.Add($labelCommand)
    $y += 1
    $textCommand = [Terminal.Gui.TextView]::new()
    $textCommand.X=2; $textCommand.Y=$y; $textCommand.Width=84; $textCommand.Height=10
    $textCommand.ReadOnly = $true
    $textCommand.Text = [NStack.ustring]::Make("(Select input/output formats to see command)")
    $dialog.Add($textCommand)
    $y += 11

    ## --- Native conversion functions ---
    function Convert-PemToDer { param($InputFile,$OutputFile) $pem=[IO.File]::ReadAllText($InputFile); $b64=($pem -replace "-----.*?-----","").Trim(); [IO.File]::WriteAllBytes($OutputFile,[Convert]::FromBase64String($b64)) }
    function Convert-DerToPem { param($InputFile,$OutputFile) $bytes=[IO.File]::ReadAllBytes($InputFile); $b64=[Convert]::ToBase64String($bytes,'InsertLineBreaks'); $pem="-----BEGIN CERTIFICATE-----`r`n$b64`r`n-----END CERTIFICATE-----"; Set-Content $OutputFile $pem -Encoding ascii }
    function Convert-PemToPkcs12 { param($CertFile,$KeyFile,$OutputFile,$Password="") $pem=Get-Content -Raw $CertFile; $b64=($pem -replace "-----.*?-----","").Trim(); $cert=[System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($b64)); $keyPem=Get-Content -Raw $KeyFile; if ($keyPem -match "PRIVATE") {$rsa=[System.Security.Cryptography.RSA]::Create(); $rsa.ImportPkcs8PrivateKey([Convert]::FromBase64String(($keyPem -replace "-----.*?-----","").Trim()),[ref]0); $cert=$cert.CopyWithPrivateKey($rsa)} else {throw "Unsupported private key format."}; [IO.File]::WriteAllBytes($OutputFile,$cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,$Password)) }
    function Convert-DerToPkcs12 { param($InputFile,$KeyFile,$OutputFile,$Password="") $tmp=[IO.Path]::GetTempFileName()+".pem"; Convert-DerToPem $InputFile $tmp; Convert-PemToPkcs12 $tmp $KeyFile $OutputFile $Password; Remove-Item $tmp -Force }

    ## --- Update command preview ---
function Update-ConversionCommand {
    $inputFile = $textInputFile.Text.ToString()
    $outputFile = $textOutputFile.Text.ToString()
    if ([string]::IsNullOrWhiteSpace($inputFile)) { 
        $textCommand.Text = [NStack.ustring]::Make("(Enter input file to see commands)")
        return 
    }

    $inputFmt = @("PEM","DER","PKCS12")[$radioInputFormat.SelectedItem]
    $outputFmt = @("PEM","DER","PKCS12")[$radioOutputFormat.SelectedItem]

    if ([string]::IsNullOrWhiteSpace($outputFile)) {
        $ext = switch ($outputFmt) {"PEM"{".pem"} "DER"{".der"} "PKCS12"{".pfx"}}
        $outputFile = [System.IO.Path]::ChangeExtension($inputFile,$ext)
        $textOutputFile.Text = [NStack.ustring]::Make($outputFile)
    }

    $keyFile = $textKeyFile.Text.ToString()
    $pw = $textPassword.Text.ToString()

    ## --- Native PowerShell functions for full conversion ---
    $psPemToDer = @"
function Convert-PemToDer {
    param([string]`$InputFile,[string]`$OutputFile)
    `$pem = Get-Content -Raw `$InputFile
    `$b64 = (`$pem -replace '-----BEGIN CERTIFICATE-----','' -replace '-----END CERTIFICATE-----','').Trim()
    [IO.File]::WriteAllBytes(`$OutputFile,[Convert]::FromBase64String(`$b64))
}
"@

    $psDerToPem = @"
function Convert-DerToPem {
    param([string]`$InputFile,[string]`$OutputFile)
    `$bytes = [IO.File]::ReadAllBytes(`$InputFile)
    `$b64 = [Convert]::ToBase64String(`$bytes,'InsertLineBreaks')
    `$pem = "-----BEGIN CERTIFICATE-----`r`n`$b64`r`n-----END CERTIFICATE-----"
    Set-Content `$OutputFile `$pem -Encoding ascii
}
"@

    $psPemToPfx = @"
function Convert-PemToPkcs12 {
    param([string]`$CertFile,[string]`$KeyFile,[string]`$OutputFile,[string]`$Password='')
    `$pem = Get-Content -Raw `$CertFile
    `$b64 = (`$pem -replace '-----BEGIN CERTIFICATE-----','' -replace '-----END CERTIFICATE-----','').Trim()
    `$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String(`$b64))
    `$keyPem = Get-Content -Raw `$KeyFile
    if (`$keyPem -match 'PRIVATE') {
        `$rsa = [System.Security.Cryptography.RSA]::Create()
        `$rsa.ImportPkcs8PrivateKey([Convert]::FromBase64String((`$keyPem -replace '-----.*?-----','').Trim()),[ref]0)
        `$cert = `$cert.CopyWithPrivateKey(`$rsa)
    } else { throw "Unsupported private key format." }
    [IO.File]::WriteAllBytes(`$OutputFile, `$cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,`$Password))
}
"@

    $psDerToPfx = @"
function Convert-DerToPkcs12 {
    param([string]`$InputFile,[string]`$KeyFile,[string]`$OutputFile,[string]`$Password='')
    `$tmp = [IO.Path]::GetTempFileName() + '.pem'
    Convert-DerToPem `$InputFile `$tmp
    Convert-PemToPkcs12 `$tmp `$KeyFile `$OutputFile `$Password
    Remove-Item `$tmp -Force
}
"@

    $psPfxToPem = @"
function Convert-Pkcs12ToPem {
    param([string]`$InputFile,[string]`$OutputFile,[string]`$Password='')
    `$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(`$InputFile,`$Password)
    `$b64 = [Convert]::ToBase64String(`$cert.RawData,'InsertLineBreaks')
    `$pem = "-----BEGIN CERTIFICATE-----`r`n`$b64`r`n-----END CERTIFICATE-----"
    Set-Content `$OutputFile `$pem -Encoding ascii
}
"@

    ## --- Tool-specific commands ---
    $cmdOpenSSL = ""; $cmdCertUtil = ""; $cmdPS = ""
    switch ("$inputFmt-$outputFmt") {
        "PEM-DER" {
            $cmdOpenSSL = "openssl x509 -in `"$inputFile`" -outform DER -out `"$outputFile`""
            $cmdCertUtil = if ($hasCertUtil) { "certutil -encodehex `"$inputFile`" `"$outputFile`" 4" } else { "" }
            $cmdPS = "$psPemToDer`nConvert-PemToDer `"$inputFile`" `"$outputFile`""
        }
        "DER-PEM" {
            $cmdOpenSSL = "openssl x509 -in `"$inputFile`" -inform DER -outform PEM -out `"$outputFile`""
            $cmdCertUtil = if ($hasCertUtil) { "certutil -decodehex `"$inputFile`" `"$outputFile`"" } else { "" }
            $cmdPS = "$psDerToPem`nConvert-DerToPem `"$inputFile`" `"$outputFile`""
        }
        "PEM-PKCS12" {
            $cmdOpenSSL = "openssl pkcs12 -export -in `"$inputFile`" -inkey `"$keyFile`" -out `"$outputFile`""
            $cmdCertUtil = "" # certutil cannot create PFX
            $cmdPS = "$psPemToPkcs12`nConvert-PemToPkcs12 `"$inputFile`" `"$keyFile`" `"$outputFile`" `"$pw`""
        }
        "DER-PKCS12" {
            $cmdOpenSSL = "# DER→PEM→PFX via OpenSSL"
            $cmdCertUtil = ""
            $cmdPS = "$psDerToPfx`nConvert-DerToPkcs12 `"$inputFile`" `"$keyFile`" `"$outputFile`" `"$pw`""
        }
        "PKCS12-PEM" {
            $cmdOpenSSL = "openssl pkcs12 -in `"$inputFile`" -out `"$outputFile`" -nodes"
            $cmdCertUtil = ""
            $cmdPS = "$psPfxToPem`nConvert-Pkcs12ToPem `"$inputFile`" `"$outputFile`" `"$pw`""
        }
        default { $cmdOpenSSL="(Input/output same)"; $cmdCertUtil="(Input/output same)"; $cmdPS="(Input/output same)" }
    }

    ## --- Determine selected tool ---
    $selectedTool = $toolOptions[$radioToolChoice.SelectedItem]
    $preview = ""
    if ($selectedTool -eq "OpenSSL") { $preview = "OpenSSL:`n  $cmdOpenSSL" }
    elseif ($selectedTool -eq "PowerShell") { $preview = "PowerShell:`n$cmdPS" }
    elseif ($selectedTool -eq "certutil") { $preview = "certutil:`n  $cmdCertUtil" }
    elseif ($selectedTool -eq "All") {
        $parts=@()
        if ($cmdOpenSSL) { $parts += "OpenSSL:`n  $cmdOpenSSL" }
        if ($cmdCertUtil) { $parts += "certutil:`n  $cmdCertUtil" }
        if ($cmdPS) { $parts += "PowerShell:`n$cmdPS" }
        $preview = $parts -join "`n`n"
    }

    $textCommand.Text = [NStack.ustring]::Make($preview)
    $textCommand.SetNeedsDisplay()
}


    ## --- Wire events ---
    $textInputFile.add_TextChanged({ Update-ConversionCommand })
    $textOutputFile.add_TextChanged({ Update-ConversionCommand })
    $radioInputFormat.add_SelectedItemChanged({ Update-ConversionCommand })
    $radioOutputFormat.add_SelectedItemChanged({ Update-ConversionCommand })
    $textKeyFile.add_TextChanged({ Update-ConversionCommand })
    $textPassword.add_TextChanged({ Update-ConversionCommand })
    $radioToolChoice.add_SelectedItemChanged({ Update-ConversionCommand })

    ## --- Convert & Close buttons on same line ---
    $btnConvert = [Terminal.Gui.Button]::new(2,$y,"Convert")
    $btnConvert.add_Clicked({
        try {
            $inputFile=$textInputFile.Text.ToString(); $outputFile=$textOutputFile.Text.ToString()
            $inputFmt=@("PEM","DER","PKCS12")[$radioInputFormat.SelectedItem]; $outputFmt=@("PEM","DER","PKCS12")[$radioOutputFormat.SelectedItem]
            $keyFile=$textKeyFile.Text.ToString(); $pw=$textPassword.Text.ToString()

            switch ("$inputFmt-$outputFmt") {
                "PEM-DER" { Convert-PemToDer $inputFile $outputFile }
                "DER-PEM" { Convert-DerToPem $inputFile $outputFile }
                "PEM-PKCS12" { Convert-PemToPkcs12 $inputFile $keyFile $outputFile $pw }
                "DER-PKCS12" { Convert-DerToPkcs12 $inputFile $keyFile $outputFile $pw }
                "PKCS12-PEM" { Convert-Pkcs12ToPem $inputFile $outputFile }
                default { Show-Modal "Notice" "Input/output formats identical. Nothing to do." }
            }
            Show-Modal "Success" "Conversion complete: $outputFile"
        } catch { Show-Modal "Error" "Conversion failed:`n$($_.Exception.Message)" }
    })
    $btnClose = [Terminal.Gui.Button]::new(20,$y,"Close")
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dialog.Add($btnConvert); $dialog.Add($btnClose)

    [Terminal.Gui.Application]::Run($dialog)
}

## Generating CSRs
function Generate-CertRequestFile {
    param(
        [hashtable]$CertData,
        [switch]$UseOpenSSL
    )

    Debug-Log "Generating certificate request file..."

    $SubjectParts = @()
    if ($CertData.CN) { $SubjectParts += "CN=$($CertData.CN)" }
    if ($CertData.OU) { $SubjectParts += "OU=$($CertData.OU)" }
    if ($CertData.O) { $SubjectParts += "O=$($CertData.O)" }
    if ($CertData.L) { $SubjectParts += "L=$($CertData.L)" }
    if ($CertData.S) { $SubjectParts += "S=$($CertData.S)" }
    if ($CertData.C) { $SubjectParts += "C=$($CertData.C)" }
    $SubjectString = $SubjectParts -join ","

    # --- CN Wildcard validation ---
    if ($CertData.CN.StartsWith("*")) {
        if (-not $CertData.CN.StartsWith("*.")) {
            Show-Modal "Validation Error" "Invalid CN wildcard format: $($CertData.CN)`n`nWildcard can only be at the start: *.domain.tld"
            return
        }
    }

    ## --- Determine output format ---
    $GenerateINF = $IsWindows -and (-not $UseOpenSSL)
    $GenerateOpenSSL = (-not $IsWindows) -or $UseOpenSSL

    $GenerateINF = $IsWindows -and (-not $UseOpenSSL)
    $GenerateOpenSSL = (-not $IsWindows) -or $UseOpenSSL

    Debug-Log "IsWindows: $IsWindows"
    Debug-Log "UseOpenSSL switch: $UseOpenSSL"
    Debug-Log "GenerateINF: $GenerateINF"
    Debug-Log "GenerateOpenSSL: $GenerateOpenSSL"

    if ($GenerateINF) {
        # --- Windows INF Generation ---
        $infFile = "$($CertData.OutputFile).inf"
        if (-not [System.IO.Path]::IsPathRooted($infFile)) {
            $infFile = Join-Path (Get-Location) $infFile
        }

        $CertReqINF = @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject = "$SubjectString"
KeySpec = 1
KeyLength = $($CertData.KeySize)
Exportable = True
ExportableEncrypted = True
MachineKeySet = False
ProviderName = Microsoft RSA SChannel Cryptographic Provider
RequestType = PKCS10
KeyUsage = 0xa0
FriendlyName = $($CertData.CN)

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.1
"@

        if ($CertData.SANs -and $CertData.SANs.Count -gt 0) {
            $CertReqINF = $CertReqINF + "`n`n[Extensions]`n2.5.29.17 = `"{text}`""
            foreach ($altName in $CertData.SANs) {
                $CertReqINF = $CertReqINF + "`n_continue_ = `"dns=$altName&`""
            }
        }

        Debug-Log "Writing INF file: $infFile"
        $CertReqINF | Out-File $infFile -Encoding ASCII
        Debug-Log "INF file generated successfully"
        return $infFile
    }
    elseif ($GenerateOpenSSL) {
        # --- OpenSSL CNF Generation ---
        $cnfFile = "$($CertData.OutputFile).cnf"
        if (-not [System.IO.Path]::IsPathRooted($cnfFile)) {
            $cnfFile = Join-Path (Get-Location) $cnfFile
        }

        $cnfContent = @"
[ req ]
default_bits       = $($CertData.KeySize)
prompt             = no
default_md         = sha256
distinguished_name = dn
req_extensions     = v3_req

[ dn ]
"@

        if ($CertData.CN) { $cnfContent += "CN = $($CertData.CN)`n" }
        if ($CertData.OU) { $cnfContent += "OU = $($CertData.OU)`n" }
        if ($CertData.O) { $cnfContent += "O = $($CertData.O)`n" }
        if ($CertData.L) { $cnfContent += "L = $($CertData.L)`n" }
        if ($CertData.S) { $cnfContent += "ST = $($CertData.S)`n" }
        if ($CertData.C) { $cnfContent += "C = $($CertData.C)`n" }

        $cnfContent += @"

[ v3_req ]
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
"@

        ## --- SANs ---
        if ($CertData.SANs -and $CertData.SANs.Count -gt 0) {
            $cnfContent += "subjectAltName = @alt_names`n"
            $cnfContent += "[ alt_names ]`n"
            for ($i = 0; $i -lt $CertData.SANs.Count; $i++) {
                $alt = $CertData.SANs[$i]

                # --- Wildcard validation: must be *.domain.tld or exact ---
                if ($alt -like "*.*" -and -not $alt.StartsWith("*.") -and $alt -ne $CertData.CN) {
                    Debug-Log "Warning: SAN $alt has an invalid wildcard format"
                }

                $cnfContent += "DNS.$($i+1) = $alt`n"
            }
        }

        Debug-Log "Writing OpenSSL CNF file: $cnfFile"
        $cnfContent | Out-File $cnfFile -Encoding ASCII
        Debug-Log "CNF file generated successfully"
        return $cnfFile
    }
}

## Show a modal to permit managing of SANs (Subject Altenrative Names) in SSL certs
function Show-ManageSANsDialog {
    if (-not $script:SubjectAltNames) { $script:SubjectAltNames = @() }
    
    if ($Verbose) {
        $ts = (Get-Date).ToString("HH:mm:ss")
        Write-Host "[$ts] LOG: Opening SANs dialog with $($script:SubjectAltNames.Count) existing SANs" -ForegroundColor Cyan
    }
    
    $dialog = [Terminal.Gui.Dialog]::new("Manage Subject Alternate Names (SANs)", 70, 22)
    $labelInfo = [Terminal.Gui.Label]::new(2, 1, "Add additional DNS names for this certificate:")
    $dialog.Add($labelInfo)
    
    ## Use TextView - simple and reliable
    $textView = [Terminal.Gui.TextView]::new()
    $textView.X = 2
    $textView.Y = 3
    $textView.Width = 64
    $textView.Height = 8
    $textView.ReadOnly = $true
    $dialog.Add($textView)
    
    if ($Verbose) {
        Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: TextView created" -ForegroundColor Cyan
    }
    
    $labelCount = [Terminal.Gui.Label]::new(2, 12, "Total SANs: 0")
    $dialog.Add($labelCount)
    
    ## Function to refresh the display
    function Update-SANDisplay {
        if ($Verbose) {
            Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Update-SANDisplay called - Current count: $($script:SubjectAltNames.Count)" -ForegroundColor Cyan
        }
        
        ## Build display text
        if ($script:SubjectAltNames -and $script:SubjectAltNames.Count -gt 0) {
            $displayText = ""
            for ($i = 0; $i -lt $script:SubjectAltNames.Count; $i++) {
                $displayText += "[$i] $($script:SubjectAltNames[$i])`n"
            }
            
            if ($Verbose) {
                Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Setting text to: $displayText" -ForegroundColor Yellow
            }
            
            $textView.Text = [NStack.ustring]::Make($displayText.TrimEnd())
        } else {
            if ($Verbose) {
                Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Setting text to empty message" -ForegroundColor Yellow
            }
            $textView.Text = [NStack.ustring]::Make("(No SANs configured)")
        }
        
        ## Update count
        $labelCount.Text = [NStack.ustring]::Make("Total SANs: $($script:SubjectAltNames.Count)")
        
        ## Force redraw
        $textView.SetNeedsDisplay()
        $labelCount.SetNeedsDisplay()
        $dialog.SetNeedsDisplay()
        [Terminal.Gui.Application]::Refresh()
        
        if ($Verbose) {
            Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Display refresh complete" -ForegroundColor Cyan
        }
    }
    
    ## Initial refresh
    if ($Verbose) {
        Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Performing initial refresh" -ForegroundColor Cyan
    }
    Update-SANDisplay
    
    $labelNewSAN = [Terminal.Gui.Label]::new(2, 14, "New DNS name:")
    $textNewSAN = [Terminal.Gui.TextField]::new(17, 14, 48, "")
    $dialog.Add($labelNewSAN)
    $dialog.Add($textNewSAN)
    
    ## Enter key handler
    $textNewSAN.add_KeyPress({
        param($e)
        if ($e.KeyEvent.Key -eq [Terminal.Gui.Key]::Enter) {
            $newSAN = $textNewSAN.Text.ToString().Trim()
            if ($Verbose) {
                Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Enter pressed with SAN: $newSAN" -ForegroundColor Cyan
            }
            if (-not [string]::IsNullOrWhiteSpace($newSAN)) {
                if ($script:SubjectAltNames -notcontains $newSAN) {
                    $script:SubjectAltNames += $newSAN
                    if ($Verbose) {
                        Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: SAN added to array, calling refresh" -ForegroundColor Cyan
                    }
                    Update-SANDisplay
                    $textNewSAN.Text = [NStack.ustring]::Make("")
                    $e.Handled = $true
                } else {
                    if ($Verbose) {
                        Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Duplicate SAN detected" -ForegroundColor Cyan
                    }
                    Show-Modal "Duplicate" "This DNS name is already in the list."
                    $e.Handled = $true
                }
            }
        }
    })
    
    $btnAdd = [Terminal.Gui.Button]::new(2, 16, "Add")
    $btnAdd.add_Clicked({
        $newSAN = $textNewSAN.Text.ToString().Trim()
        if ($Verbose) {
            Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Add button clicked with SAN: $newSAN" -ForegroundColor Cyan
        }
        if (-not [string]::IsNullOrWhiteSpace($newSAN)) {
            if ($script:SubjectAltNames -notcontains $newSAN) {
                $script:SubjectAltNames += $newSAN
                if ($Verbose) {
                    Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: SAN added to array, calling refresh" -ForegroundColor Cyan
                }
                Update-SANDisplay
                $textNewSAN.Text = [NStack.ustring]::Make("")
            } else {
                if ($Verbose) {
                    Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Duplicate SAN detected" -ForegroundColor Cyan
                }
                Show-Modal "Duplicate" "This DNS name is already in the list."
            }
        }
    })
    $dialog.Add($btnAdd)
    
    $btnRemove = [Terminal.Gui.Button]::new(12, 16, "Remove by #")
    $btnRemove.add_Clicked({
        if ($script:SubjectAltNames.Count -eq 0) {
            Show-Modal "No SANs" "No SANs to remove"
            return
        }
        
        $removeDialog = [Terminal.Gui.Dialog]::new("Remove SAN", 40, 10)
        $removeDialog.Add([Terminal.Gui.Label]::new(2, 1, "Enter SAN number to remove:"))
        $textNum = [Terminal.Gui.TextField]::new(2, 3, 10, "")
        $removeDialog.Add($textNum)
        
        $btnOK = [Terminal.Gui.Button]::new(2, 5, "OK")
        $btnOK.add_Clicked({
            $numText = $textNum.Text.ToString().Trim()
            if ($numText -match '^\d+$') {
                $idx = [int]$numText
                if ($idx -ge 0 -and $idx -lt $script:SubjectAltNames.Count) {
                    if ($Verbose) {
                        Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Removing SAN at index $idx" -ForegroundColor Cyan
                    }
                    $script:SubjectAltNames = @($script:SubjectAltNames | Select-Object -Index (0..($script:SubjectAltNames.Count-1) | Where-Object { $_ -ne $idx }))
                    [Terminal.Gui.Application]::RequestStop()
                } else {
                    Show-Modal "Invalid" "Number out of range"
                }
            }
        })
        $removeDialog.Add($btnOK)
        
        $btnCancelRemove = [Terminal.Gui.Button]::new(10, 5, "Cancel")
        $btnCancelRemove.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
        $removeDialog.Add($btnCancelRemove)
        
        [Terminal.Gui.Application]::Run($removeDialog)
        Update-SANDisplay
    })
    $dialog.Add($btnRemove)
    
    $btnClear = [Terminal.Gui.Button]::new(32, 16, "Clear All")
    $btnClear.add_Clicked({
        if ($script:SubjectAltNames.Count -eq 0) { return }
        $result = [Terminal.Gui.MessageBox]::Query("Confirm", "Clear all SANs?", @("Yes", "No"))
        if ($result -eq 0) {
            if ($Verbose) {
                Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Clearing all SANs" -ForegroundColor Cyan
            }
            $script:SubjectAltNames = @()
            Update-SANDisplay
        }
    })
    $dialog.Add($btnClear)
    
    $btnClose = [Terminal.Gui.Button]::new(46, 16, "Close")
    $btnClose.add_Clicked({ 
        if ($Verbose) {
            Write-Host "[$([DateTime]::Now.ToString('HH:mm:ss'))] LOG: Closing SANs dialog with $($script:SubjectAltNames.Count) SANs" -ForegroundColor Cyan
        }
        [Terminal.Gui.Application]::RequestStop() 
    })
    $dialog.Add($btnClose)
    
    $hintLabel = [Terminal.Gui.Label]::new(2, 19, "Tip: Press Enter after typing to quickly add a SAN")
    $dialog.Add($hintLabel)
    
    [Terminal.Gui.Application]::Run($dialog)
}

## Create Certificate window
function Show-NewCertificateDialog {
    param([hashtable]$PrefilledData = $null)
    
    $dialog = [Terminal.Gui.Dialog]::new("New Certificate Request", 80, 32)
    $y = 1
    
    $labelCN = [Terminal.Gui.Label]::new(2, $y, "Common Name (CN):")
    $textCN = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.CN ?? ""))
    $dialog.Add($labelCN); $dialog.Add($textCN); $y += 2
    
    $labelOrg = [Terminal.Gui.Label]::new(2, $y, "Organization (O):")
    $textOrg = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.O ?? ""))
    $dialog.Add($labelOrg); $dialog.Add($textOrg); $y += 2
    
    $labelOU = [Terminal.Gui.Label]::new(2, $y, "Org Unit (OU):")
    $textOU = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.OU ?? ""))
    $dialog.Add($labelOU); $dialog.Add($textOU); $y += 2
    
    $labelCity = [Terminal.Gui.Label]::new(2, $y, "City/Locality (L):")
    $textCity = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.L ?? ""))
    $dialog.Add($labelCity); $dialog.Add($textCity); $y += 2
    
    $labelState = [Terminal.Gui.Label]::new(2, $y, "State/Province (S):")
    $textState = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.S ?? ""))
    $dialog.Add($labelState); $dialog.Add($textState); $y += 2
    
    $labelCountry = [Terminal.Gui.Label]::new(2, $y, "Country Code (C):")
    $textCountry = [Terminal.Gui.TextField]::new(25, $y, 3, ($PrefilledData.C ?? ""))
    $btnValidateCountry = [Terminal.Gui.Button]::new(30, $y, "?")
    $btnValidateCountry.add_Clicked({
        $code = $textCountry.Text.ToString().Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($code)) {
            Show-Modal "Country Code" "Enter a 2-letter ISO country code (e.g., US, GB, CA, DE)`n`nLeave empty if not required."
            return
        }

        $validCodes = Get-ValidCountryCodes

        ## Check for common mistakes first
        $commonMistakes = @{
            "UK" = "GB"  # United Kingdom
            "EN" = "GB"  # England
            "SW" = "SE"  # Sweden (not Switzerland)
            "SZ" = "CH"  # Switzerland
            "NL" = "NL"  # Netherlands (this one is actually correct, but people doubt it)
        }

        if ($commonMistakes.ContainsKey($code)) {
            $correctCode = $commonMistakes[$code]
            $mistakeMsg = "Did you mean '$correctCode'?`n`n'$code' is not a valid ISO country code.`n`nCommon mistakes:`n- UK → GB (United Kingdom)`n- EN → GB (England)`n- SW → SE (Sweden) or CH (Switzerland)`n`nThe correct code for your country might be '$correctCode'.`n`nSee: https://knowledge.digicert.com/general-information/ssl-certificate-country-codes"
            Show-Modal "Country Code Hint" $mistakeMsg -EnableCopy
        }
        elseif ($validCodes -contains $code) {
            $validMsg = "Country code " + $code + " is valid!"
            Show-Modal "Valid" $validMsg
        } else {
            $invalidMsg = "Country code " + $code + " is not valid`n`nExamples: US, GB, CA, DE, FR, JP, AU`n`nSee: https://knowledge.digicert.com/general-information/ssl-certificate-country-codes"
            Show-Modal "Invalid" $invalidMsg -EnableCopy
        }
    })
    $dialog.Add($labelCountry); $dialog.Add($textCountry); $dialog.Add($btnValidateCountry); $y += 2
    
    $labelEmail = [Terminal.Gui.Label]::new(2, $y, "Email Address:")
    $textEmail = [Terminal.Gui.TextField]::new(25, $y, 45, ($PrefilledData.Email ?? ""))
    # Ensure the button is on the same line as the email TextField
    $textEmail = [Terminal.Gui.TextField]::new(25, $y, 45, ($PrefilledData.Email ?? ""))
    $btnValidateEmail = [Terminal.Gui.Button]::new(71, $y, "?")  # same $y as textEmail

    $dialog.Add($labelEmail); $dialog.Add($textEmail); $y += 2
    
    $btnValidateEmail.add_Clicked({
        $email = $textEmail.Text.ToString().Trim()
        $cnText = $textCN.Text.ToString().Trim()

        if (-not $email) {
            Show-Modal "Email Validation" "Enter an email address first."
            return
        }

        ## Check basic email format
        if ($email -notmatch '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') {
            Show-Modal "Email Validation" "Invalid email format."
            return
        }

        ## Check domain against CN (wildcards allowed)
        $emailDomain = $email -replace '^.*@', ''
        if ($cnText -notlike "*$emailDomain*") {
            Show-Modal "Email Validation" "Email domain '$emailDomain' does not match CN '$cnText'."
        } else {
            Show-Modal "Email Validation" "Email address is valid."
        }
    })
    $dialog.Add($btnValidateEmail)

    
    $labelKeySize = [Terminal.Gui.Label]::new(2, $y, "Key Size (bits):")
    $keySizes = @("2048", "3072", "4096", "8192")
    $radioKeySize = [Terminal.Gui.RadioGroup]::new(25, $y, $keySizes)
    $radioKeySize.SelectedItem = if ($PrefilledData.KeySizeIndex) { $PrefilledData.KeySizeIndex } else { 0 }
    $dialog.Add($labelKeySize); $dialog.Add($radioKeySize); $y += 5
    
    $labelSANs = [Terminal.Gui.Label]::new(2, $y, "Subject Alt Names:")
    $sanCount = if ($script:SubjectAltNames) { $script:SubjectAltNames.Count } else { 0 }
    $labelSANCount = [Terminal.Gui.Label]::new(25, $y, "($sanCount configured - F3 to manage)")
    $dialog.Add($labelSANs); $dialog.Add($labelSANCount); $y += 2
    
    $labelOutput = [Terminal.Gui.Label]::new(2, $y, "Output Filename:")
    $textOutput = [Terminal.Gui.TextField]::new(25, $y, 50, ($PrefilledData.OutputFile ?? "certificate"))
    $dialog.Add($labelOutput); $dialog.Add($textOutput); $y += 2
    
    $btnManageSANs = [Terminal.Gui.Button]::new(2, $y + 1, "Manage SANs (F3)")
    $btnManageSANs.add_Clicked({
        $currentData = @{
            CN = $textCN.Text.ToString()
            O = $textOrg.Text.ToString()
            OU = $textOU.Text.ToString()
            L = $textCity.Text.ToString()
            S = $textState.Text.ToString()
            C = $textCountry.Text.ToString()
            Email = $textEmail.Text.ToString()
            KeySizeIndex = $radioKeySize.SelectedItem
            OutputFile = $textOutput.Text.ToString()
        }
        
        [Terminal.Gui.Application]::RequestStop()
        Show-ManageSANsDialog
        Show-NewCertificateDialog -PrefilledData $currentData
    })
    $dialog.Add($btnManageSANs)
    
    $btnReviewSANs = [Terminal.Gui.Button]::new(22, $y + 1, "Review SANs")
    $btnReviewSANs.add_Clicked({
        if ($script:SubjectAltNames -and $script:SubjectAltNames.Count -gt 0) {
            $sanList = $script:SubjectAltNames -join "`n- "
            Show-Modal "Configured SANs" ("The following SANs are configured:`n`n- " + $sanList)
        } else {
            Show-Modal "No SANs" "No Subject Alternate Names configured."
        }
    })
    $dialog.Add($btnReviewSANs)
    
    $btnGenerateINF = [Terminal.Gui.Button]::new(2, $y + 3, "Generate INF")
    $btnGenerateINF.add_Clicked({
        if ([string]::IsNullOrWhiteSpace($textCN.Text.ToString())) { 
            Show-Modal "Validation Error" "Common Name (CN) is required!"
            return
        }
        
        $countryText = $textCountry.Text.ToString().Trim().ToUpper()
        $emailText = $textEmail.Text.ToString().Trim()
        $cnText = $textCN.Text.ToString().Trim()
        
        if (-not [string]::IsNullOrWhiteSpace($countryText)) {
            $validCodes = Get-ValidCountryCodes
            if ($validCodes -notcontains $countryText) {
                $countryMsg = "Country code " + $countryText + " is not valid.`n`nContinue anyway?"
                $result = [Terminal.Gui.MessageBox]::Query("Invalid Country Code", $countryMsg, @("Yes", "No"))
                if ($result -ne 0) { return }
            }
        }
        
        if (-not (Validate-EmailAgainstCN -Email $emailText -CN $cnText)) {
            Show-Modal "Validation Error" "Email domain does not match the CN (wildcards allowed).`n`nPlease correct the email before generating CSR or INF."
            return
        }
        
        $certData = @{
            CN = $cnText
            O = $textOrg.Text.ToString()
            OU = $textOU.Text.ToString()
            L = $textCity.Text.ToString()
            S = $textState.Text.ToString()
            C = $countryText
            Email = $emailText
            KeySize = $keySizes[$radioKeySize.SelectedItem]
            OutputFile = $textOutput.Text.ToString()
            SANs = $script:SubjectAltNames
        }
        
        try { 
            $infFile = Generate-CertRequestFile -CertData $certData
            $outFile = $certData.OutputFile + ".csr"
            $successMsg = "INF file generated successfully`n`nFile created: " + $infFile + "`n`nYou can now run:`ncertreq.exe -new " + $infFile + " " + $outFile
            Show-Modal "Success" $successMsg
            # [Terminal.Gui.Application]::RequestStop() ## Uncomment if you want the modal ot close immediately after creating the file
        } catch { 
            Show-Modal "Error" ("Failed to generate INF file:`n" + $_.Exception.Message)
        }
    })
    $dialog.Add($btnGenerateINF)

    ## --- Generate OpenSSL CNF Button ---
    $btnGenerateCNF = [Terminal.Gui.Button]::new(18, $y + 3, "Generate CNF (OpenSSL)")
    $btnGenerateCNF.add_Clicked({
    if ([string]::IsNullOrWhiteSpace($textCN.Text.ToString())) {
      Show-Modal "Validation Error" "Common Name (CN) is required!"
      return
    }

    $countryText = $textCountry.Text.ToString().Trim().ToUpper()
    $emailText = $textEmail.Text.ToString().Trim()
    $cnText = $textCN.Text.ToString().Trim()

    if (-not [string]::IsNullOrWhiteSpace($countryText)) {
        $validCodes = Get-ValidCountryCodes
        if ($validCodes -notcontains $countryText) {
            $countryMsg = "Country code $countryText is not valid.`n`nContinue anyway?"
            $result = [Terminal.Gui.MessageBox]::Query("Invalid Country Code", $countryMsg, @("Yes", "No"))
            if ($result -ne 0) { return }
        }
    }

    if (-not (Validate-EmailAgainstCN -Email $emailText -CN $cnText)) {
        Show-Modal "Validation Error" "Email domain does not match the CN (wildcards allowed).`n`nPlease correct the email before generating CNF."
        return
    }

$certData = @{
    CN = $cnText
    O = $textOrg.Text.ToString()
    OU = $textOU.Text.ToString()
    L = $textCity.Text.ToString()
    S = $textState.Text.ToString()
    C = $countryText
    Email = $emailText
    KeySize = $keySizes[$radioKeySize.SelectedItem]
    OutputFile = $textOutput.Text.ToString()
    SANs = $script:SubjectAltNames
}

    try {
        $cnfFile = Generate-CertRequestFile -CertData $certData -UseOpenSSL
        $successMsg = "OpenSSL CNF file generated successfully!`n`nFile created: $cnfFile`n`nYou can now run:`nopenssl req -new -config $cnfFile -keyout <keyfile> -out <csrfile>"
        Show-Modal "Success" $successMsg
#        [Terminal.Gui.Application]::RequestStop() ## uncomment this if you want the dialog to exit immediately after creation of file
    } catch {
        Show-Modal "Error" ("Failed to generate CNF file:`n" + $_.Exception.Message)
    }
    })
    $dialog.Add($btnGenerateCNF)
    
    $btnGenerate = [Terminal.Gui.Button]::new(44, $y + 3, "Generate CSR")
    $btnGenerate.add_Clicked({
        if ([string]::IsNullOrWhiteSpace($textCN.Text.ToString())) { 
            Show-Modal "Validation Error" "Common Name (CN) is required!"
            return
        }
        
        $countryText = $textCountry.Text.ToString().Trim().ToUpper()
        $emailText = $textEmail.Text.ToString().Trim()

        $emailText = $textEmail.Text.ToString().Trim()
        $cnText = $textCN.Text.ToString().Trim()

        # Validate email format
        if ($emailText -and $emailText -notmatch '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') {
            Show-Modal "Validation Error" "Email address is not valid."
            return
        }

        # Validate domain matches CN
        if ($emailText -and $emailText -match '@(.+)$') {
            $emailDomain = $matches[1]
            if ($cnText -notlike "*$emailDomain*") {
                Show-Modal "Validation Error" "Email domain '$emailDomain' does not match CN '$cnText'."
                return
            }
        }

        $cnText = $textCN.Text.ToString().Trim()
        
        if (-not [string]::IsNullOrWhiteSpace($countryText)) {
            $validCodes = Get-ValidCountryCodes
            if ($validCodes -notcontains $countryText) {
                $countryMsg = "Country code " + $countryText + " is not valid.`n`nContinue anyway?"
                $result = [Terminal.Gui.MessageBox]::Query("Invalid Country Code", $countryMsg, @("Yes", "No"))
                if ($result -ne 0) { return }
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($emailText)) {
            if ($emailText -match '@(.+)$') {
                $emailDomain = $matches[1]
                if ($cnText -notlike "*$emailDomain*") {
                    $warnMsg = "Warning: Email domain " + $emailDomain + " does not match CN " + $cnText + ".`n`nContinue anyway?"
                    $result = [Terminal.Gui.MessageBox]::Query("Domain Mismatch", $warnMsg, @("Yes", "No"))
                    if ($result -ne 0) { return }
                }
            }
        }
        
        $certData = @{
            CN = $cnText
            O = $textOrg.Text.ToString()
            OU = $textOU.Text.ToString()
            L = $textCity.Text.ToString()
            S = $textState.Text.ToString()
            C = $countryText
            Email = $emailText
            KeySize = $keySizes[$radioKeySize.SelectedItem]
            OutputFile = $textOutput.Text.ToString()
            SANs = $script:SubjectAltNames
}
        
        if ($script:SubjectAltNames -and $script:SubjectAltNames.Count -gt 0) {
            $sanList = $script:SubjectAltNames -join "`n- "
            $sanMsg = "The following SANs will be included:`n`n- " + $sanList + "`n`nProceed?"
            $result = [Terminal.Gui.MessageBox]::Query("Confirm SANs", $sanMsg, @("Yes", "No"))
            if ($result -ne 0) { return }
        }
        
        try { 
            Generate-CertificateRequest -CertData $certData
            $outFile = $certData.OutputFile + ".csr"
            $successMsg = "Certificate request generated successfully!`n`nFile created: " + $outFile + "`n`nThe private key is stored in your Windows Certificate Store (Current User).`nUse certmgr.msc to manage it."
            Show-Modal "Success" $successMsg
            [Terminal.Gui.Application]::RequestStop()
        } catch { 
            Show-Modal "Error" ("Failed to generate certificate:`n" + $_.Exception.Message)
        }
    })
    $dialog.Add($btnGenerate)
    
    $btnCancel = [Terminal.Gui.Button]::new(60, $y + 3, "Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dialog.Add($btnCancel)
    
    [Terminal.Gui.Application]::Run($dialog)
}

## File Browser
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

## Show info and helper sites about a certificate the user gives us
function Show-CertInfoDialog {
    $dialog = [Terminal.Gui.Dialog]::new("Open Certificate File", 70, 12)
    
    $labelInfo = [Terminal.Gui.Label]::new(2, 1, "Enter path to certificate file (.crt, .cer, .csr, .p7b):")
    $dialog.Add($labelInfo)
    
    $textPath = [Terminal.Gui.TextField]::new(2, 3, 64, "")
    $dialog.Add($textPath)
    
$btnBrowse = [Terminal.Gui.Button]::new(2, 5, "Browse...")
$btnBrowse.add_Clicked({
    $selectedFile = Show-FileBrowserDialog -StartDir "." -Title "Select Certificate File" -Filter @("*.crt", "*.cer", "*.pem", "*.*")
    
    if ($selectedFile) {
        $textPath.Text = [NStack.ustring]::Make($selectedFile)
        Debug-Log "Selected file: $selectedFile"
    }
})
$dialog.Add($btnBrowse)

$btnOpen = [Terminal.Gui.Button]::new(18, 5, "Open")
$btnOpen.add_Clicked({
    $filePath = $textPath.Text.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        Show-Modal "Error" "Please enter a file path"
        return
    }
    
    if (-not (Test-Path $filePath)) {
        Show-Modal "Error" ("File not found: " + $filePath)
        return
    }
    
    [Terminal.Gui.Application]::RequestStop()
    
    try {
        $certInfo = Get-CertificateInfo -FilePath $filePath
        Show-Modal "Certificate Information" $certInfo
    } catch {
        Show-Modal "Error" ("Failed to read certificate:`n" + $_.Exception.Message)
    }
})
    $dialog.Add($btnOpen)   
    $btnCancel = [Terminal.Gui.Button]::new(28, 5, "Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dialog.Add($btnCancel)
    
    [Terminal.Gui.Application]::Run($dialog)
}

## Read certificate info and act accordingly
function Get-CertificateInfo {
    param([string]$FilePath)

    Debug-Log ("Reading certificate file: $FilePath")

    if (-not (Test-Path $FilePath)) { throw "File not found: $FilePath" }

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $fileText  = Get-Content -Raw $FilePath

    ## Detect PEM formats
    $isPemCert = $fileText -match '-----BEGIN CERTIFICATE-----'
    $isPemCSR  = $fileText -match '-----BEGIN CERTIFICATE REQUEST-----'

    ## Read platform detection variables (do NOT overwrite)
    $platformWindows = $IsWindows
    $platformLinux   = $IsLinux
    $platformMacOS   = $IsMacOS

    ## Determine if OpenSSL is available
    $hasOpenSSL = (Get-Command "openssl" -ErrorAction SilentlyContinue) -ne $null

    ## Windows has certutil by default, but prefer OpenSSL if available
    $useCertUtil = $platformWindows -and (-not $hasOpenSSL)

    ## --- Helper for running external tools ---
    function Run-Tool($cmd, $args) {
        $output = & $cmd $args 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            throw "External tool failed:`n$cmd $args`n$output"
        }
        return $output
    }

    ## CSR HANDLING
    if ($extension -eq ".csr" -or $isPemCSR) {

        # Convert PEM CSR → DER temp file
        if ($isPemCSR) {
            if ($fileText -match '-----BEGIN CERTIFICATE REQUEST-----(.+?)-----END CERTIFICATE REQUEST-----') {
                $b64 = ($matches[1] -replace '\s','')
                $derBytes = [Convert]::FromBase64String($b64)
                $tmpFile = [System.IO.Path]::GetTempFileName()
                [System.IO.File]::WriteAllBytes($tmpFile, $derBytes)
                $FilePath = $tmpFile
            } else {
                throw "Invalid PEM CSR block"
            }
        }

        ## Parse using certutil or openssl
        if ($useCertUtil) {
            $output = Run-Tool "certutil.exe" "-dump `"$FilePath`""
        } else {
            if (-not $hasOpenSSL) {
                throw "OpenSSL is required on Linux/macOS. Install: apt install openssl OR brew install openssl"
            }
            $output = Run-Tool "openssl" "req -in `"$FilePath`" -noout -text"
        }

        $info = "=== Certificate Signing Request ===`n`n"

        ## Subject
        if ($output -match 'Subject:([^\r\n]+)') {
            $info += "Subject: $($matches[1].Trim())`n"
        }

        ## Key size (OpenSSL format)
        if ($output -match 'Public-Key: \((\d+) bit\)') {
            $info += "Key Size: $($matches[1]) bits`n"
        }
        # certutil format
        elseif ($output -match 'Public Key Length:\s*(\d+)\s*bits') {
            $info += "Key Size: $($matches[1]) bits`n"
        }

        ## SANs
        $sans = @()
        $output -split "`n" | ForEach-Object {
            if ($_ -match 'DNS Name=(.+)') { $sans += $matches[1].Trim() }
            elseif ($_ -match 'DNS:([^\s,]+)') { $sans += $matches[1].Trim() }
        }
        if ($sans.Count -gt 0) {
            $info += "`nSubject Alternative Names:`n"
            $sans | ForEach-Object { $info += "  - $_`n" }
        }

        $info += "`n--- Raw Output ---`n$output"
        return $info
    }

    ## CERTIFICATE HANDLING
    elseif ($extension -in ".crt", ".cer", ".p7b", ".pem" -or $isPemCert) {

        # Convert PEM → DER
        if ($isPemCert) {
            if ($fileText -match '-----BEGIN CERTIFICATE-----(.+?)-----END CERTIFICATE-----') {
                $b64 = ($matches[1] -replace '\s','')
                $bytes = [Convert]::FromBase64String($b64)
                $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes)
            } else {
                throw "Invalid PEM certificate block"
            }
        } else {
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($FilePath)
        }

        $info  = "=== Certificate Information ===`n`n"
        $info += "Subject: $($cert.Subject)`n"
        $info += "Issuer: $($cert.Issuer)`n"
        $info += "Valid From: $($cert.NotBefore)`n"
        $info += "Valid To: $($cert.NotAfter)`n"
        $info += "Serial Number: $($cert.SerialNumber)`n"
        $info += "Thumbprint: $($cert.Thumbprint)`n"

        if ($cert.PublicKey.Key -and $cert.PublicKey.Key.KeySize) {
            $info += "Key Size: $($cert.PublicKey.Key.KeySize) bits`n"
        }

        # SAN extraction
        $sanExt = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
        if ($sanExt) {
            $info += "`nSubject Alternate Names:`n"
            $sanExt.Format($false) -split ', ' | ForEach-Object {
                $info += "  - $_`n"
            }
        }

        $info += "`nSHA Thumbprints: https://cert.sh`n"
        $info += "Chain Builder: https://whatsmychaincert.com/"

        return $info
    }

    else {
        throw "Unsupported file type: $extension. Supported: .csr, .crt, .cer, .p7b, .pem"
    }
}

## Main UI Setup - Having defined the functions, start the main body of this script

## Select theme before proceeding
$themes = Get-Theme -mode $Theme

## Apply global theme
[Terminal.Gui.Colors]::Base      = $themes.Global
[Terminal.Gui.Colors]::Dialog    = $themes.Global
[Terminal.Gui.Colors]::Menu      = $themes.Global
#[Terminal.Gui.Colors]::TopLevel = $themes.Global
#[Terminal.Gui.Colors]::Base     = $themes.Global

## Create main window with its special scheme
$win = [Terminal.Gui.Window]::new(" PS SSL Helper")
$win.ColorScheme = $themes.MainWindow

$script:SubjectAltNames = @()
$script:LastGeneratedCSR = $null

Debug-Log "=== PS SSL Helper Starting ==="

$win = [Terminal.Gui.Window]::new(" PS SSL Helper ${BuildVersion} - Jordbaer Build")
$win.X = 0
$win.Y = 1
$win.Width = [Terminal.Gui.Dim]::Fill()
$win.Height = [Terminal.Gui.Dim]::Fill()

$welcomeLabel = [Terminal.Gui.Label]::new(2, 2, @"
Welcome to PS SSL Helper ${BuildVersion} (Jordbaer Build)

Press F1 for help
Press F2 to create a new certificate
Press F3 to manage SANs
Press F4 to convert certificates
Press F5 for to parse an exsiting certificate's info
Press F10 to quit
"@)

$win.Add($welcomeLabel)

## Status bar
$statusBar = [Terminal.Gui.StatusBar]::new(@(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1, "~F1~ Help", {
        Show-Modal "Shortcuts" "F1 - Help`nF2 - New Cert`nF3 - SANs`nF4 Convert`nF5 Cert Info`nF10 - Quit" 
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F2, "~F2~ New Cert", {
        Show-NewCertificateDialog
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F3, "~F3~ SANs", {
        Show-ManageSANsDialog
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F4, "~F4~ Cert Info", {
        Show-CertInfoDialog
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10, "~F10~ Quit", {
        [Terminal.Gui.Application]::RequestStop()
    })
))

## Menus
$menuFile = [Terminal.Gui.MenuBarItem]::new("_File", @(
    [Terminal.Gui.MenuItem]::new("E_xit (F10)", "Quit", {
        [Terminal.Gui.Application]::RequestStop() 
    })
))

$menuActions = [Terminal.Gui.MenuBarItem]::new("_Actions", @(
    [Terminal.Gui.MenuItem]::new("_New Certificate (F2)", "Create New Certificate", {
        Show-NewCertificateDialog
    }),
    [Terminal.Gui.MenuItem]::new("_Manage SANs (F3)", "Subject Alternate Names", {
        Show-ManageSANsDialog
    }),
    [Terminal.Gui.MenuItem]::new("_Convert (F4)", "Convert Certificate", {
        Show-DownloadExampleCertDialog
    }),
    [Terminal.Gui.MenuItem]::new("Cert _Info (F5)", "Show Certificate Info", {
        Show-CertInfoDialog
    })
))

$menuHelp = [Terminal.Gui.MenuBarItem]::new("_Help", @(
    [Terminal.Gui.MenuItem]::new("_Keys (F1)", "Shortcuts", { 
        Show-Modal "Shortcuts" "F1 - Help`nF2 - New Cert`nF3 - Manage SANs`nF4 Convert`nF5 Certificate Info`nF10 - Quit" 
    }),
    [Terminal.Gui.MenuItem]::new("_About", "About", { 
        Show-Modal "About" "PS SSL Helper ${BuildVersion}`nGPL-3 Copyleft`nBy Knightmare2600`nhttps://github.com/knightmare2600" -EnableCopy
    }),
    [Terminal.Gui.MenuItem]::new("Why _Jordbaer", "Why the codename?", {
        Show-JordbaerInfo
    })
))

$menu = [Terminal.Gui.MenuBar]::new(@($menuFile, $menuActions, $menuHelp))

## Get the theme
$themes = Get-Theme -mode $Theme

## Apply theme to the main application root (TopLevel)
if ($TopLevel.PSObject.Properties.Name -contains "ColorScheme") {
    $TopLevel.ColorScheme = $themes.Global
}

## Apply theme to main window
if ($MainWindow.PSObject.Properties.Name -contains "ColorScheme") {
    $MainWindow.ColorScheme = $themes.MainWindow
}

## MenuBar also supports ColorScheme
if ($Menu.PSObject.Properties.Name -contains "ColorScheme") {
    $Menu.ColorScheme = $themes.Global
}

## Same for StatusBar
if ($StatusBar.PSObject.Properties.Name -contains "ColorScheme") {
    $StatusBar.ColorScheme = $themes.Global
}

## Add everything to the application
[Terminal.Gui.Application]::Top.Add($menu)
[Terminal.Gui.Application]::Top.Add($win)
[Terminal.Gui.Application]::Top.Add($statusBar)

Debug-Log "=== PS SSL Helper Ready ==="
Debug-Log "File | Actions | Help"

## Run the application
try {
    [Terminal.Gui.Application]::Run()
} finally {
    [Terminal.Gui.Application]::Shutdown()
}
