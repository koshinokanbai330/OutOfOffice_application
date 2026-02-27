# OutOfOffice_application

Outlook VSTO Add-in for Windows Classic Outlook (Microsoft 365 Apps) that lets you:

- Create **all-day meeting requests** (Business Trip / Full Day Off / AM Half Day Off / PM Half Day Off) as Outlook appointments and send them.
- Persist **To/Cc recipients** to `%USERPROFILE%\Documents\mailingList.txt` and reload them on next launch.
- Configure **OOF (auto-reply)** for the absence period via Microsoft Graph API (delegated `MailboxSettings.ReadWrite`), including the user's Outlook HTML signature.
- Automatically **download the travel-allowance Excel template** and fill in trip data (date, destination, departure/return times) when the leave type is Business Trip.

## Requirements

| Component | Version |
|-----------|---------|
| Windows Classic Outlook | Microsoft 365 Apps (Version 2502+) 64-bit |
| Visual Studio | 2022 (Community/Professional/Enterprise) |
| .NET Framework | 4.8 |
| Visual Studio Workloads | *Office/SharePoint Development* |

## Build Instructions

1. **Clone** this repository.

2. **Open** `OutOfOfficeAddin.sln` in Visual Studio 2022.

3. **Restore NuGet packages** – Visual Studio restores them automatically on first build,
   or run:
   ```
   nuget restore OutOfOfficeAddin.sln
   ```

4. **Configure Azure AD** (required for OOF auto-reply) – see [Graph Setup](#microsoft-graph-setup) below.

5. **Build** the solution (`Ctrl+Shift+B` or *Build → Build Solution*).

6. **Register the add-in** – on first build Visual Studio registers the add-in in the
   current user's registry (`HKCU\...\Outlook\Addins\OutOfOfficeAddin`).
   If you want to deploy to another machine, use the ClickOnce publish output from
   *Build → Publish*.

## Microsoft Graph Setup

The OOF (auto-reply) feature requires a delegated Microsoft Graph token with the
`MailboxSettings.ReadWrite` scope.

### Register an app in Azure Active Directory

1. Go to **Azure Portal** → **Azure Active Directory** → **App registrations** → **New registration**.
2. Give it a name (e.g., `OutOfOffice VSTO Addin`).
3. Set **Supported account types** to *Accounts in this organizational directory only* (or your
   appropriate choice).
4. Set **Redirect URI** → *Public client/native (mobile & desktop)* → `http://localhost`.
5. Click **Register**.
6. Note the **Application (client) ID** and **Directory (tenant) ID**.
7. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** →
   search for `MailboxSettings.ReadWrite` → Add. Grant admin consent if required.

### Configure `appsettings.json`

Edit `OutOfOfficeAddin/appsettings.json` and replace the placeholder values:

```json
{
  "AzureAd": {
    "ClientId": "<paste Application (client) ID>",
    "TenantId": "<paste Directory (tenant) ID>",
    "RedirectUri": "http://localhost"
  }
}
```

> **Note**: Never commit real secrets to source control.  
> The `appsettings.json` file is included only as a template; add it to `.gitignore` after
> populating it with real values.

## Usage

After building and launching Outlook you will see an **Out of Office** task pane on the right.

### Fields

| Field | Description |
|-------|-------------|
| **Type** | Business Trip / Full Day Off / AM Half Day Off / PM Half Day Off |
| **Start date / End date** | Date range of absence |
| **Subject (auto)** | Auto-generated: `{FamilyName} BT`, `{FamilyName} OFF`, etc. (read-only) |
| **Location** | Defaults to *Home* for Off types, empty for Business Trip |
| **To / Cc** | Semicolon-separated addresses; or use **Add from Address Book** |
| **Set automatic replies** | Enables OOF auto-reply via Graph (default ON) |
| **Internal / External message preview** | Read-only preview of the auto-reply text |
| **Create and fill allowance Excel** | *(Business Trip only)* Downloads template and fills data |
| **Excel save folder** | *(Business Trip only)* Folder where the Excel file is saved |

### Buttons

| Button | Description |
|--------|-------------|
| **Create draft** | Saves the meeting as a draft (To not required) |
| **Send** | Sends the meeting, sets OOF, and optionally creates Excel (To required; folder required when Excel is enabled) |
| **Cancel** | Resets the form |

### What happens on Send

1. An all-day Outlook meeting is created (Show as: Free, Reminder: Off) and sent to To/Cc recipients.
2. To/Cc are saved to `%USERPROFILE%\Documents\mailingList.txt`.
3. If **Set automatic replies** is checked, OOF is configured via Graph for the selected period.
   - Back-date shown in messages = **End date + 1 day**.
   - The user's Outlook HTML signature (from `%APPDATA%\Microsoft\Signatures\`) is appended.
4. If **Business Trip** and **Create and fill allowance Excel** is checked:
   - Template downloaded from the Bosch Confluence URL.
   - Filled with dates, destination, departure/arrival times.
   - Saved as `BT-Allowance-{FamilyName}-{yyyyMMdd}.xlsx` in the chosen folder.
5. The status log at the bottom shows what completed and any errors.

If OOF or Excel steps fail after the meeting was sent, a warning is displayed and the log
shows which steps succeeded and which failed.

## Subject Rules

| Leave type | Subject |
|------------|---------|
| Business Trip | `{FamilyName} BT` |
| Full Day Off | `{FamilyName} OFF` |
| AM Half Day Off | `{FamilyName} AM OFF` |
| PM Half Day Off | `{FamilyName} PM OFF` |

## Excel Template Sheets

| Trip duration | Sheets filled |
|---------------|---------------|
| 1 day | `日帰り One-Day` only |
| 2+ days | `日帰り One-Day` **and** `宿泊 Overnight` |

Times written: Departure 07:00 · Start 09:00 · Finish 18:00 · Return 21:00.

## Project Structure

```
OutOfOfficeAddin.sln
OutOfOfficeAddin/
  OutOfOfficeAddin.csproj        – VSTO project (net48)
  ThisAddIn.cs / .Designer.cs    – Add-in entry point
  appsettings.json               – Azure AD config template
  Models/
    LeaveType.cs                 – Leave type enum
    OutOfOfficeRequest.cs        – Request data model
  Services/
    SubjectHelper.cs             – Meeting subject / default location
    MailingListService.cs        – Persist To/Cc to mailingList.txt
    SignatureService.cs          – Read Outlook HTML signature
    GraphAuthService.cs          – MSAL authentication helper
    OofService.cs                – Set OOF via Microsoft Graph
    MeetingService.cs            – Create/send Outlook meeting
    ExcelService.cs              – Download + fill allowance Excel
  UI/
    RelayCommand.cs              – ICommand helper
    TaskPaneViewModel.cs         – WPF ViewModel
    TaskPaneView.xaml / .cs      – WPF UserControl (task pane UI)
    TaskPaneHost.cs              – WinForms wrapper for VSTO task pane
  Properties/
    AssemblyInfo.cs
```
