
SharePoint Upload (Microsoft Login + File Picker v8 + OBO)

A minimal, production-friendly starter that adds **Microsoft (SharePoint/OneDrive) upload** as the 3rd option on your AgentFleet Upload page.

- **Frontend (React + MSAL)** — One-click Microsoft login (Auth Code + PKCE), launch **File Picker v8** for SharePoint/OneDrive, send selected items to backend.
- **Backend (Node + Express + MSAL-Node)** — **On-Behalf-Of (OBO)** exchange to call Microsoft Graph and **download** selected files.
- **Ingestion Mock** — Simple endpoint to receive downloaded files, so you can swap with your real AgentFleet ingestion.

## High-Level Flow

1) User clicks **Sign in with Microsoft** on the upload page (MSAL).
2) User clicks **Pick from SharePoint/OneDrive** — opens File Picker v8.
3) Picker returns `{ driveId, itemId }` for each selected file (delegated access).
4) Frontend POSTs `{ driveId, itemId }` and user token to backend.
5) Backend performs **OBO** to obtain a Graph token and **streams file bytes** from Graph.
6) File is forwarded to your ingestion service (replace `ingestion-mock`).

## Prerequisites

- Node 18+
- Azure Entra ID (Azure AD) tenant
- Two App Registrations:
  - **SPA App (Frontend)**: used by MSAL in the browser
  - **Backend App (Confidential)**: used by Express server for OBO

### Permissions (delegated)

Grant admin consent (tenant or per-customer tenant) to the SPA app for scopes you'll request:
- `openid profile email offline_access`
- `Files.Read` (or `Files.Read.All` for cross-site breadth)
- (Optional) `Sites.Read.All` if you want SharePoint site-browsing helpers

> The backend app needs no extra Graph app permissions for pure OBO — it redeems the user's delegated token to call Graph on their behalf.

## Quick Start

### 1) Configure Frontend

Copy `.env.example` to `.env` and set:

```
VITE_AZURE_AD_TENANT_ID=<your-tenant-id or 'common'>
VITE_AZURE_AD_CLIENT_ID=<your-spa-app-client-id>
VITE_GRAPH_SCOPES=openid profile email offline_access Files.Read
VITE_PICKER_BASEURL_SHAREPOINT=https://<yourtenant>.sharepoint.com
VITE_PICKER_BASEURL_ONEDRIVE=https://<yourtenant>-my.sharepoint.com
VITE_BACKEND_URL=http://localhost:4000
```

Then:

```bash
cd frontend
npm install
npm run dev
```

### 2) Configure Backend

Copy `.env.example` to `.env` and set:

```
TENANT_ID=<your-tenant-id>
BACKEND_CLIENT_ID=<your-backend-confidential-app-client-id>
BACKEND_CLIENT_SECRET=<your-backend-client-secret>
BACKEND_REDIRECT_URI=http://localhost:4000/auth/redirect   # used for OBO caches
INGEST_URL=http://localhost:5001/ingest
PORT=4000
```

Then:

```bash
cd backend
npm install
npm run dev
```

### 3) Start Ingestion Mock (optional)

```bash
cd ingestion-mock
npm install
npm start
```

### 4) Try It

- Open **http://localhost:5173** (Vite dev server).
- Click **Sign in with Microsoft**.
- Click **Pick from SharePoint** (or OneDrive) — select one or more files.
- Click **Upload Selected** to send the selection to backend; backend downloads via Graph and forwards to ingestion.

## Notes

- The included File Picker code path uses the **v8 POST to /_layouts/15/FilePicker.aspx** with `postMessage` channel to receive picks. You must set the **base URL** appropriately per tenant (examples in `.env.example`). 
- For multi-tenant SaaS, set SPA app as multi-tenant and handle per-tenant admin consent.
- Replace `ingestion-mock` with your real AgentFleet ingestion endpoint.

## Security

- SPA uses **Auth Code + PKCE** and stores tokens in session storage.
- Backend keeps OBO tokens server-side (not exposed to the browser).
- All SharePoint/OneDrive listing is **delegated** — users only see what they have access to.
