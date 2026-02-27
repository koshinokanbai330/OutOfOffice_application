# Out of Office Add-in (Office.js / React / TypeScript)

An Outlook task-pane add-in that automates out-of-office workflows:

- Creates calendar events (all-day, with attendees)
- Sets automatic reply (OOF) via Microsoft Graph
- Generates and fills a travel-allowance Excel file in OneDrive (Business Trip mode)
- Persists mailing lists to OneDrive (App Folder)

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | 18 or later |
| npm | 9 or later |
| Microsoft 365 subscription | Any plan with Exchange/Outlook |
| Azure AD app registration | See below |

---

## Azure AD App Registration

1. Open [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**.
2. **Name**: `OutOfOffice Addin`  
   **Supported account types**: *Accounts in this organizational directory only* (or multi-tenant as needed)  
   **Redirect URI**: `Single-page application (SPA)` → `https://localhost:3000`
3. After creation, note the **Application (client) ID** and **Directory (tenant) ID**.
4. Under **API permissions**, add the following **Delegated** Microsoft Graph permissions and grant admin consent:
   - `User.Read`
   - `MailboxSettings.ReadWrite`
   - `Calendars.ReadWrite`
   - `Files.ReadWrite`
   - `offline_access`
   - `openid`
   - `profile`
5. Under **Expose an API**, set the Application ID URI to `api://localhost:3000/<CLIENT_ID>` and add a scope named `access_as_user`.

---

## Environment Setup

```bash
cd office-addin
cp .env.example .env
```

Edit `.env`:

```
REACT_APP_CLIENT_ID=<your-application-client-id>
REACT_APP_TENANT_ID=<your-directory-tenant-id>
```

---

## Build & Run

```bash
cd office-addin
npm install
npm run build        # production build → dist/
npm start            # dev server at https://localhost:3000 (self-signed cert)
```

> **Note:** The dev server uses a self-signed HTTPS certificate. Trust it in your browser the first time.

---

## Excel Template

Place (or replace) `office-addin/assets/template.xlsx` with your actual travel-allowance template.  
The file **must** contain two worksheets:

- `日帰り One-Day` — for day-trip entries
- `宿泊 Overnight` — for multi-day overnight entries

Each sheet should have a header row containing `日付` in the first cell of the date column. The add-in writes rows starting immediately below that header.

---

## Sideloading the Add-in

### Outlook on Windows

1. Build or start the dev server so `https://localhost:3000/taskpane.html` is reachable.
2. Open Outlook → **File** → **Manage Add-ins** → **My add-ins** → **Add a custom add-in** → **Add from File…**
3. Select `office-addin/manifest.xml`.

### Outlook on Mac

1. Ensure the dev server is running.
2. Open Outlook → **Tools** → **Add-ins…** → **+** (bottom-left) → **Add from file…**
3. Select `office-addin/manifest.xml`.

### Outlook on the Web (OWA)

1. Open [outlook.office.com](https://outlook.office.com) → **Settings** → **View all Outlook settings** → **Mail** → **Customize actions**.
2. Or navigate to: `https://outlook.office.com/mail/inclientstore`.
3. Click **+** → **Add from file** and upload `manifest.xml`.

---

## Centralized Admin Deployment

To deploy to all users in your Microsoft 365 tenant:

1. Open [admin.microsoft.com](https://admin.microsoft.com) → **Settings** → **Integrated apps** → **Upload custom apps**.
2. Choose **Office Add-in** and upload `manifest.xml`.
3. Assign to users or groups and click **Deploy**.

---

## Usage Guide

Open any email or appointment in Outlook and click **Out of Office** in the ribbon.

| Field | Description |
|---|---|
| **Type** | Business Trip, Full Day Off, AM Half Day Off, PM Half Day Off |
| **Start / End date** | Inclusive date range for the calendar event and OOF period |
| **Subject (auto)** | Auto-generated from your display name surname + type suffix |
| **Location** | Pre-filled with "Home" for off-type; blank for Business Trip |
| **To / Cc** | Semicolon-separated recipient emails; saved to OneDrive for reuse |
| **Set automatic replies** | Toggles OOF scheduling via Graph API |
| **Internal / External preview** | Live preview of the generated OOF HTML messages |
| **Signature** | Optional HTML appended to both OOF messages |
| **Destination** | (Business Trip only) Written into the Excel allowance rows |
| **Create draft** | Creates calendar event without sending invites or setting OOF |
| **Send** | Creates event with invites, sets OOF, and (if BT) creates Excel |
| **Cancel** | Resets all fields to defaults |

---

## Project Structure

```
office-addin/
├── assets/
│   └── template.xlsx          # Excel template (replace with real file)
├── src/
│   ├── index.html             # HTML entry point
│   ├── index.tsx              # Office.onReady bootstrap
│   └── taskpane/
│       ├── App.tsx            # Root React component
│       ├── components/
│       │   ├── TaskPane.tsx   # Main UI component
│       │   └── StatusLog.tsx  # Operation log display
│       ├── hooks/
│       │   ├── useAuth.ts     # MS Graph profile fetch
│       │   └── useMailingList.ts
│       └── services/
│           ├── authService.ts       # Office SSO + MSAL fallback
│           ├── calendarService.ts   # Graph calendar events
│           ├── excelService.ts      # OneDrive Excel template copy & fill
│           ├── mailingListService.ts
│           ├── oofService.ts        # Graph mailboxSettings OOF
│           └── subjectHelper.ts    # Date & subject utilities
├── manifest.xml               # Office Add-in manifest
├── package.json
├── tsconfig.json
└── webpack.config.js
```
