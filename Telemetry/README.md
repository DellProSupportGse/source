# Telemetry Tool Manager

**Version:** 4.2 (GUI title)  
**Platform:** Windows PowerShell 5.1+ / Windows Forms  
**Purpose:** Register Microsoft GSE PowerShell tools for centralized Azure telemetry reporting and manage their dashboard visibility without modifying individual tools or deleting telemetry.

> **Version note:** The script's opening comment identifies `v4.1`, while the application window identifies `v4.2`. This README follows the visible GUI version. Align both strings in the script for consistent version reporting.

## Overview

Telemetry Tool Manager provides one place to onboard tools into the Microsoft GSE telemetry system. It:

1. Connects to the designated Microsoft GSE Azure tenant and subscription.
2. Creates or verifies a telemetry-specific Azure Table Storage table for each tool.
3. Creates or updates that tool's entry in the central `TelemetryToolsConfig` registry.
4. Generates PowerShell telemetry code to embed in the tool.
5. Lists registered tools and allows dashboard visibility to be enabled or disabled without deleting collected data or stopping ingestion.

The manager is a **configuration and onboarding utility**, not a telemetry collector itself. Once a tool incorporates the generated code, **that tool** sends telemetry to the `PostTelemetryData` Azure Function when it starts.

## How the components fit together

```text
                  TELEMETRY TOOL MANAGER
                  (Windows Forms / PowerShell)
                              |
                   Sign in to Microsoft GSE
                              |
                   Azure Table Storage
                              |
                +-------------+--------------+
                |                            |
        TelemetryToolsConfig         <Tool>TelemetryData
         (central tool registry)     (per-tool event records)
                ^                            ^
                |                            |
         Create/update entry          PostTelemetryData
                |                       Azure Function
                |                            ^
                |                            |
         Manager GUI                  PowerShell tool
                                     (generated snippet)

       Registry-driven Azure Functions / dashboard
                    |
       Read enabled tools and available summaries
                    |
             Telemetry dashboard
```

The wider telemetry solution also uses registry-driven functions such as `GetTelemetryToolsConfig`, `GetTelemetrySummary`, `GetTelemetrySummaryFast`, and `UpdateTelemetrySummaryData`. **Those functions and the dashboard are separate components**; their implementations are not included in this script. Their precise behavior should be verified against their deployed versions.

## Azure configuration

The current script is configured with:

| Setting | Value |
| --- | --- |
| Storage account | `gsetools` |
| Resource group | `GSETools` |
| Azure region indicated by Function URL | Central US |
| Registry table | `TelemetryToolsConfig` |
| Registry partition key | `Tool` |
| Telemetry table naming convention | `<ToolName>TelemetryData` |
| Azure Function | `PostTelemetryData` |
| Tenant and subscription | Fixed in the script to the Microsoft GSE environment |
| Telemetry Function URL | `https://gsetools-bufhdqefb8e6ecc6.centralus-01.azurewebsites.net/api/PostTelemetryData` |

The script hardcodes its environment identifiers; it does not offer a multi-tenant or multi-subscription selector.

## Requirements and access

- Windows with **Windows PowerShell 5.1 or later** and Windows Forms support. The program is designed for Windows desktop use, not PowerShell Core on Linux/macOS.
- Network access to Azure authentication, Azure Resource Manager, Azure Table Storage, and the telemetry Function endpoint.
- `Az.Accounts` and `Az.Storage` PowerShell modules.
- A Microsoft GSE identity authorized for the configured tenant/subscription, with permission to **list storage account keys** for the `gsetools` storage account.

The manager checks for the required modules and attempts to install missing modules from the PowerShell Gallery into the **CurrentUser** scope; elevation is not intended to be required. At shutdown, it attempts to unload modules it imported and uninstall only the two modules it installed during that run. Existing module installations are not deliberately removed.

> **Security:** The manager retrieves a storage account key to authenticate direct Azure Table Storage requests using SharedKeyLite. It does not intentionally print or export the key. However, the key grants broad storage account access and must be protected. This is **not** a least-privilege table-data RBAC design. Restrict who can list storage keys.

## Quick start

1. Run the PowerShell GUI script in a Windows PowerShell session.
2. Select **Connect to Microsoft GSE** and complete Microsoft authentication if prompted.
3. Enter the new tool's name, estimated minutes saved, and the PowerShell expression holding its version.
4. Choose whether to show the tool in the dashboard and include it in aggregate totals.
5. Select **Preview Code** to inspect the generated integration snippet.
6. Select **Create / Update Tool**, review the proposed changes, and confirm.
7. After the GUI reports **SAVED TO AZURE**, select **Copy Code**.
8. Paste the generated snippet into the tool's PowerShell source **after the tool's version variable is initialized**.
9. Run the instrumented tool. Verify its output reports successful telemetry submission.
10. Run the telemetry summary updater in the deployed Azure environment. The tool should appear in the dashboard once telemetry has been received and the required summaries have been refreshed.

> Saving a registration does **not** automatically edit the tool's script, execute that tool, submit an initial event, or run the summary updater.

## User interface

### Connect to Microsoft GSE

The **Connect to Microsoft GSE** button:

- Initializes the required Azure PowerShell modules.
- Checks whether an existing Azure context matches the configured tenant and subscription.
- Signs in, if necessary, and selects the configured subscription.
- Finds the `gsetools` storage account in `GSETools`.
- Requests an account key and makes it available to the GUI for subsequent storage operations.

Azure sign-in and key retrieval run in a background PowerShell runspace so the Windows Forms interface remains responsive. The GUI reports a warning after **60 seconds** if a connection has not completed, but does not forcibly terminate the worker at that point. The Connect button remains disabled until that operation finishes.

A displayed **CONNECTED** state confirms that the manager obtained a storage key and initialized a storage context. It does **not** mean every Azure Table operation has already been tested; the script checks tables when saving or managing tools.

### New tool registration fields

| Field | Meaning |
| --- | --- |
| **Tool name** | Identifier for the tool. Must start with a letter and contain 3–41 ASCII letters/digits. Whitespace is automatically removed. |
| **Azure telemetry table** | Read-only name generated as `<ToolName>TelemetryData`. |
| **Minutes saved per run** | Nonnegative integer from `0` through `100000` used as the registered estimate of time savings. |
| **Version expression** | PowerShell variable, property, or hashtable reference evaluated when the embedded telemetry code runs. |
| **Show in telemetry dashboard** | Saves the registry's `Enabled` flag. |
| **Include in All Tools totals** | Saves the registry's `IncludeInTotals` flag for consumers of the registry. |

Valid version expressions include:

```powershell
$Ver
$script:Version
$Config.Version
$Config['Version']
$Config["Version"]
$Config['Build'].Version
```

The manager validates the **syntax** of the expression, not whether it resolves to a valid, nonempty value at runtime. Make sure the referenced version variable or field exists **before** the telemetry call.

### Buttons

- **Preview Code** — Validate current inputs, update the predicted table name, and generate the snippet locally. Does not require saving to Azure.
- **Create / Update Tool** — Confirm and apply Azure changes, then verify the registry record. Requires a completed connection.
- **Copy Code** — Copy the displayed PowerShell snippet to the clipboard. Does not install it in another script.
- **Manage Existing Tools** — List registry entries and toggle their dashboard visibility.

The activity pane provides timestamps and status messages for module setup, Azure access, registration, and errors.

## What happens when a tool is registered

Assume a tool is named `ECHO`.

### 1. Verify or create the tables

The manager checks for:

```text
TelemetryToolsConfig
ECHOTelemetryData
```

If a table does not exist, the manager creates it. An already existing table is reused; registering an existing tool does not remove historical telemetry.

### 2. Create or update its registry entity

The central registry entry uses:

```text
Table:        TelemetryToolsConfig
PartitionKey: Tool
RowKey:       ECHO
```

Example entity:

```json
{
  "PartitionKey": "Tool",
  "RowKey": "ECHO",
  "TableName": "ECHOTelemetryData",
  "MinutesSaved": 60,
  "Enabled": true,
  "IncludeInTotals": true
}
```

If the entity already exists, the manager updates it with the submitted fields. If no entity exists, it creates one. The save worker then reads the record back and verifies that its `TableName` matches the expected table.

**Important:** The version expression is used to generate the PowerShell snippet; it is **not stored** in the registry entity by this script.

### 3. Generate integration code

The preview substitutes the tool-specific table name and version expression into a common telemetry implementation. Registering a tool and embedding that implementation are two separate steps.

## What the generated telemetry code does

### Startup initialization

The snippet creates a per-invocation report ID, initializes geolocation state, enables telemetry, and includes a flag intended to prevent multiple startup events from being sent in the same process:

```powershell
$script:TelemetryReportID    = [guid]::NewGuid().Guid
$script:TelemetryGeoResolved = $false
$script:TelemetryGeoData     = @{}
$script:TelemetryStartupSent = $false
$script:uploadToAzure        = $true
```

### Machine identifier

`Get-TelemetryMachineHash` combines the Windows domain, username, and computer name, computes a SHA-256 digest, and sends the **first 24 hexadecimal characters** of that digest.

The original identifiers are not sent directly in the generated telemetry payload. However, the result is a **stable pseudonymous identifier, not anonymization**: it can link events from the same environment, and its source components may be guessable.

### Geolocation

`Resolve-TelemetryGeo` attempts the following sequence:

1. Read the country and country code from **Windows Regional Settings**, and the time zone from the operating system.
2. If local region resolution fails, query `https://ipwho.is/` for country, country code, and time zone.
3. If both methods fail, preserve the available local time zone and leave unknown country fields empty.

Geolocation is cached for the lifetime of the initialized snippet. The script prints the resolved country and source for visibility.

**Privacy/network note:** Local Windows regional settings represent configured locale, not necessarily the computer's actual physical location. The `ipwho.is` fallback makes an external request that may expose the public source IP address to that service, even though the generated Azure telemetry data does not explicitly include an IP address.

### Build the telemetry event

The `Send-ToolTelemetry` function builds a data record with these fields:

| Field | Value |
| --- | --- |
| `PartitionKey` | Tool telemetry name with a trailing `TelemetryData` suffix removed |
| `RowKey` | New GUID for the telemetry event |
| `PSVersion` | Running PowerShell version |
| `countryCode` | Resolved two-letter country code, if available |
| `MachineHash` | Truncated SHA-256 machine/user-derived hash |
| `Version` | Value of the supplied version expression |
| `timezone` | Resolved time zone |
| `ReportID` | GUID created when the snippet is initialized |
| `country` | Resolved country name |

It submits a JSON envelope with:

```json
{
  "TelemetryName": "ECHOTelemetryData",
  "TableName": "ECHOTelemetryData",
  "Data": {
    "PartitionKey": "ECHO",
    "RowKey": "<new-guid>",
    "PSVersion": "<powershell-version>",
    "countryCode": "US",
    "MachineHash": "<24-character-hash>",
    "Version": "<tool-version>",
    "timezone": "<local-timezone-id>",
    "ReportID": "<report-guid>",
    "country": "United States"
  }
}
```

The bracketed values above are illustrative placeholders, **not actual emitted values**.

Although `Send-ToolTelemetry` also accepts `EventName`, `ServerCount`, and `GroupCount`, the generated `Data` object **does not currently include those values**. `EventName` is used for startup-event deduplication rather than being transmitted.

### Submit to Azure

The generated code sends the envelope using HTTP POST to:

```text
https://gsetools-bufhdqefb8e6ecc6.centralus-01.azurewebsites.net/api/PostTelemetryData
```

The request uses JSON and a 15-second HTTP timeout. The intended downstream flow is:

```text
Tool starts
   |
Resolve location and build telemetry event
   |
POST to PostTelemetryData
   |
Azure Function receives/validates request
   |
Tool-specific Azure Table Storage table
   |
Summary updater processes stored events
   |
Registry-driven dashboard shows updated metrics
```

The **Azure Function's server-side validation, authentication, and write behavior are not contained in this manager script**, so consult its deployed source when documenting those details.

The generated code considers an HTTP request that completes without a terminating exception a success. It does not independently query Azure Table Storage to verify the stored row.

### Telemetry does not stop the tool

If telemetry submission fails, the exception is caught and displayed as a warning. The parent application should continue running.

Example console output:

```text
  Resolving Geo Location...
    Country: United States
    Source : Windows Regional Settings
  Recording telemetry...
  Telemetry recorded successfully
```

For troubleshooting, set `DebugTelemetry = $true` in the generated call to display the request body and returned response. Be careful when sharing debug output because it contains event identifiers and the pseudonymous machine hash.

## Managing existing tools

Select **Manage Existing Tools** after connecting.

The dialog loads entities from `TelemetryToolsConfig` with `PartitionKey = 'Tool'` and displays:

- Tool name
- Dashboard visibility (`Enabled` / `Disabled`)
- Telemetry table
- Minutes saved

Select a tool, then select **Enable / Disable Selected**. The manager confirms the choice and makes an Azure Table Storage **MERGE** request that changes **only** the `Enabled` property.

| Action | Dashboard visibility | Telemetry ingestion | Historical telemetry |
| --- | --- | --- | --- |
| Enable tool | Intended to show tool | Continues | Preserved |
| Disable tool | Intended to hide tool | Continues | Preserved |
| Change `IncludeInTotals` through registration | Separate registry setting | Continues | Preserved |
| Register/update tool | Controlled by chosen flags | Unchanged until code is embedded and run | Preserved |

The manager does not delete tool tables or telemetry entities when disabling a tool. Changes may not immediately appear in dashboard summaries until the deployed registry/dashboard consumers refresh their data.

**Current limitation:** The Manage Existing Tools dialog toggles `Enabled` only. To change `IncludeInTotals` or `MinutesSaved`, use the main registration fields and **Create / Update Tool** with the intended complete values. Be careful not to overwrite existing settings unintentionally.

## Background operation and responsiveness

The current implementation uses separate PowerShell runspaces for:

- Azure connection and key retrieval
- Table creation and tool registration
- Listing and toggling existing tool entries

The UI uses Windows Forms timers to observe their completion and update the status displays. This design is intended to prevent the main form from freezing during network requests.

- Connection requests display a timeout warning at 60 seconds.
- Registration displays a warning after 90 seconds.
- Neither warning guarantees the background worker has been canceled.
- A management operation must finish before that management dialog can close.

## Common issues and troubleshooting

| Symptom | Likely cause / next step |
| --- | --- |
| Cannot connect | Verify Microsoft GSE access, authentication, configured subscription, connectivity, and permission to list the storage account keys. |
| Manager appears connected but save fails | A connection does not validate all table operations. Check the activity log for the HTTP status and failed table request. |
| `403` from storage | Verify that the key is valid, account networking permits access, and the account permits shared-key authorization. |
| Tool name rejected | Use 3–41 letters/digits, starting with a letter. Spaces are removed automatically; punctuation is not supported. |
| Version parameter is empty | Initialize the referenced version variable before telemetry execution and ensure its value is nonempty. |
| Registration saved but tool absent from dashboard | Embed the code, execute the tool, confirm event ingestion, run the summary updater, and ensure `Enabled` is true. |
| Telemetry says failed | Check network connectivity, the Function endpoint, its access requirements, and Azure Function logs. Telemetry exceptions are intentionally nonfatal for the tool. |
| Tool is hidden but telemetry still arrives | Expected behavior. `Enabled` controls dashboard/summary visibility, not data collection. |
| Telemetry location appears wrong | Windows regional settings may differ from physical location. |

## Operational and security considerations

1. **Protect storage account keys.** Users able to list account keys may have much broader permissions than are required for editing this one registry.
2. **Do not assume the telemetry Function is authenticated.** The generated snippet does not pass an API key or bearer token; review the Function's deployed authentication policy and endpoint exposure.
3. **Respect privacy expectations.** The machine hash is pseudonymous and the fallback geolocation service receives a network request.
4. **Review generated code before embedding it.** The snippet contains function declarations and script-scoped variables; check for naming conflicts in existing tools.
5. **Place the call after version initialization.** The manager's input validation does not execute the version expression.
6. **Understand the event deduplication behavior.** The startup guard becomes true before the HTTP request completes. A failed startup POST will not automatically be retried by another startup invocation in that same initialized session.
7. **Check summary refresh behavior.** Registration alone does not populate telemetry history or refresh reporting data.
8. **Treat success reporting narrowly.** The tool's success message confirms a non-error response from the HTTP endpoint, not independent storage verification.
9. **Plan for upgrades.** Updating a registry record does not automatically update snippets already copied into other PowerShell tools.

## Scope and limitations

This utility currently supports registering tools, generating a standardized telemetry snippet, and managing the `Enabled` flag in the central registry.

It does **not**:

- Edit existing PowerShell tool files automatically.
- Delete or purge telemetry records.
- Stop telemetry transmission from a tool when dashboard visibility is disabled.
- Change deployed Azure Function code or dashboard code.
- Automatically run the summary updater.
- Provide the complete server-side data model or dashboard calculation logic.
- Verify the deployed `PostTelemetryData` Function's implementation beyond the response to the generated HTTP request.

## Suggested validation checklist

After onboarding a new tool:

- [ ] Connection shows the correct Microsoft GSE environment.
- [ ] The correct `<ToolName>TelemetryData` table exists.
- [ ] The entry exists in `TelemetryToolsConfig` with the intended flags.
- [ ] Generated code is placed after initialization of the version variable.
- [ ] Running the tool prints `Telemetry recorded successfully`.
- [ ] A new telemetry row can be located in the tool-specific table.
- [ ] The deployed summary updater completes successfully.
- [ ] The dashboard reflects the new tool and expected aggregate inclusion.
- [ ] Disabling the tool hides it from the dashboard while subsequent telemetry events are still stored.

---

**Maintainer note:** This README describes behavior found in the supplied Telemetry Tool Manager script. End-to-end behaviors involving Azure Functions, scheduled summary updates, and dashboard rendering should be confirmed against those separately deployed components.
