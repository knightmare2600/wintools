# Creating Demo Data for DSA TUI

This guide explains how to create your own demo data files for testing and development with DSA TUI.

## File Format: JSONC (JSON with Comments)

DSA TUI supports **JSONC** files (`.jsonc` extension) which allow you to add comments to standard JSON.

### Comment Styles Supported

```jsonc
## This is a comment using double hash
{
  "users": [
    ## Comments can appear above objects
    {
      "Name": "John Smith",  ## Or inline after values
      "Email": "john@example.com"
    }
  ]
}
```

---

## File Structure

Your `.jsonc` file should have this top-level structure:

```jsonc
{
  "users": [ ... ],
  "groups": [ ... ],
  "computers": [ ... ],
  "domainControllers": [ ... ]
}
```

---

## User Objects

Each user should have these properties:

```jsonc
{
  "Name": "Annie Lennox",
  "SamAccountName": "annie.lennox",
  "UserPrincipalName": "annie.lennox@example.org",
  "OU": ["Locations", "UK", "Scotland", "Aberdeen", "Eurythmics"],
  "Groups": ["Eurythmics", "Vocalists", "Keyboardists"],
  "Title": "Lead Vocalist/Keyboardist",
  "Email": "annie.lennox@example.org",
  "Country": "UK",
  "Disabled": false,
  "Locked": false,
  "MustChangePassword": false,
  "Department": "Music",
  "Office": "Aberdeen Office",
  "Phone": "+44 1224 496 010",
  "MobilePhone": "+44 7700 941001",
  "Street": "210 Union Street",
  "City": "Aberdeen",
  "PostalCode": "AB10 1TL",
  "Company": "Example Music Ltd",
  "Manager": "Dave Stewart",
  "Description": "Lead vocalist and co-founder"
}
```

### Required Fields
- `Name`
- `SamAccountName`
- `UserPrincipalName`
- `Email`

### Optional Fields
All other fields are optional and will use defaults if not specified.

---

## Group Objects

```jsonc
{
  "Name": "Simple Minds",
  "Description": "Scottish rock band formed in Glasgow in 1977",
  "Type": "Security",
  "Scope": "Global",
  "ManagedBy": "Jim Kerr",
  "Email": "simpleminds@example.com"
}
```

### Required Fields
- `Name`
- `Type` (values: `"Security"` or `"Distribution"`)
- `Scope` (values: `"Global"`, `"Universal"`, or `"DomainLocal"`)

---

## Computer Objects

```jsonc
{
  "Name": "EXAMBPABD001",
  "SamAccountName": "EXAMBPABD001$",
  "Type": "Macbook",
  "Role": "MBP",
  "Site": "ABD",
  "Location": "Aberdeen, UK",
  "DNSHostName": "EXAMBPABD001.example.org",
  "OU": ["Locations", "UK", "Scotland", "Aberdeen", "Computers"],
  "OS": "macOS Sonoma",
  "OperatingSystemVersion": "14.2.1",
  "Description": "MacBook – Annie Lennox",
  "IPv4Address": "192.168.224.137",
  "Domain": "example.org",
  "Enabled": true,
  "TPMEnabled": false,
  "TPMActivated": false
}
```

### Required Fields
- `Name`
- `SamAccountName` (must end with `$`)
- `DNSHostName`
- `Type`

---

## Domain Controller Objects

```jsonc
{
  "Name": "EXAGLADC01",
  "SamAccountName": "EXAGLADC01$",
  "DNSHostName": "EXAGLADC01.example.com",
  "HostName": "EXAGLADC01.example.com",
  "Site": "GLA",
  "Location": "Glasgow, Scotland",
  "Domain": "example.com",
  "Forest": "jukebox.example",
  "OS": "Windows Server 2022 Standard",
  "OperatingSystemVersion": "10.0 (20348)",
  "IPv4Address": "192.168.4.20",
  "Enabled": true,
  "IsGlobalCatalog": true,
  "IsReadOnly": false,
  "FSMORoles": ["Schema Master", "Domain Naming Master", "PDC Emulator"],
  "ReplicationHealth": "Healthy"
}
```

### Required Fields
- `Name`
- `DNSHostName`
- `Domain`
- `Site`

---

## Complete Example File

Here's a minimal complete example:

**my-demo-data.jsonc**
```jsonc
{
  ## ===== Users =====
  "users": [
    {
      "Name": "John Smith",
      "SamAccountName": "john.smith",
      "UserPrincipalName": "john.smith@example.com",
      "Email": "john.smith@example.com",
      "OU": ["Users", "IT"],
      "Groups": ["IT Staff", "VPN Users"],
      "Title": "System Administrator",
      "Department": "IT",
      "Disabled": false,
      "Locked": false
    },
    {
      "Name": "Jane Doe",
      "SamAccountName": "jane.doe",
      "UserPrincipalName": "jane.doe@example.com",
      "Email": "jane.doe@example.com",
      "OU": ["Users", "Sales"],
      "Groups": ["Sales Team"],
      "Title": "Sales Manager",
      "Department": "Sales",
      "Disabled": false,
      "Locked": false
    }
  ],

  ## ===== Groups =====
  "groups": [
    {
      "Name": "IT Staff",
      "Description": "IT Department Staff",
      "Type": "Security",
      "Scope": "Global",
      "Email": "it@example.com"
    },
    {
      "Name": "Sales Team",
      "Description": "Sales Department",
      "Type": "Security",
      "Scope": "Global",
      "Email": "sales@example.com"
    }
  ],

  ## ===== Computers =====
  "computers": [
    {
      "Name": "DESK-IT-001",
      "SamAccountName": "DESK-IT-001$",
      "DNSHostName": "DESK-IT-001.example.com",
      "Type": "Workstation",
      "OU": ["Computers", "IT"],
      "OS": "Windows 11 Pro",
      "IPv4Address": "192.168.1.100",
      "Domain": "example.com",
      "Enabled": true
    }
  ],

  ## ===== Domain Controllers =====
  "domainControllers": [
    {
      "Name": "DC01",
      "SamAccountName": "DC01$",
      "DNSHostName": "DC01.example.com",
      "HostName": "DC01.example.com",
      "Site": "Default-First-Site-Name",
      "Location": "Main Office",
      "Domain": "example.com",
      "Forest": "example.com",
      "OS": "Windows Server 2022",
      "IPv4Address": "192.168.1.10",
      "Enabled": true,
      "IsGlobalCatalog": true,
      "IsReadOnly": false,
      "FSMORoles": ["Schema Master", "PDC Emulator"],
      "ReplicationHealth": "Healthy"
    }
  ]
}
```

---

## Loading Your Demo Data

1. Save your file with `.jsonc` extension (e.g., `my-demo-data.jsonc`)
2. Launch DSA TUI
3. Go to **File → Load Demo Data** 
4. Select your `.jsonc` file
5. The data will be imported and displayed in the tree view

---

## Tips & Best Practices

### 1. Use Comments Liberally
```jsonc
## ===== Marketing Department =====
## Located in Building A, Floor 3
{
  "Name": "Marketing Team",
  "Description": "Marketing and Creative",
  "Type": "Security",
  "Scope": "Global"
}
```

### 2. Organize by Department or Location
Group users/computers by OU structure to match real Active Directory

### 3. Test with Small Files First
Start with 5-10 users to test your structure before creating hundreds

### 4. Validate JSON Structure
Use an online JSON validator if you get errors (just remove comments first)

### 5. Boolean Values
Use lowercase: `true` and `false` (not `True`/`False` or `"true"`)

### 6. Arrays
Always use square brackets for lists:
```jsonc
"Groups": ["Group1", "Group2", "Group3"]
"OU": ["Top", "Level", "Path"]
```

### 7. Null Values
Use `null` for empty/missing values:
```jsonc
"Manager": null,
"IPv6Address": null
```

---

## Common Mistakes

❌ **Wrong:**
```jsonc
{
  "Name": "John Smith",
  "Disabled": True,  ## Capital T
  "Groups": "IT Staff",  ## String instead of array
  ## Missing comma here
  "Email": "john@example.com"
}
```

✅ **Correct:**
```jsonc
{
  "Name": "John Smith",
  "Disabled": true,
  "Groups": ["IT Staff"],
  "Email": "john@example.com"
}
```

---

## Need Help?

- Check the included `demo-data-example.jsonc` for a complete working example
- Validate your JSON structure at jsonlint.com (remove `##` comments first)
- Contact support if you encounter issues

---

**Happy Testing! 🎵**
