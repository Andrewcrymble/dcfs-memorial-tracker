# Outlook Integration Setup

The DC&S Memorial Tracker can show your Microsoft 365 inbox on the dashboard
and surface customer-related emails on each order. To enable it, you need to
register a free Azure AD application (one-time, ~5 minutes) and paste its
Client ID into `index.html`.

## Step 1 — Register the Azure AD application

1. Go to <https://entra.microsoft.com/> and sign in with the Microsoft 365
   account you use for **info@crymbleandsons.com**.
2. In the left sidebar, click **Identity → Applications → App registrations**.
3. Click **+ New registration** at the top.
4. Fill in:
   - **Name**: `DC&S Memorial Tracker`
   - **Supported account types**: *Accounts in any organizational directory
     and personal Microsoft accounts* (the default if you're unsure)
   - **Redirect URI**:
     - Platform dropdown: **Single-page application (SPA)**
     - URL: `https://tracker.crymbleandsons.com/`
       (include the trailing slash — must exactly match where the tracker
       is served from)
5. Click **Register**.

You'll land on the application's overview page.

## Step 2 — Copy the Client ID

On the overview page, find **Application (client) ID** — a long string like
`12345678-1234-1234-1234-123456789abc`. Copy it.

## Step 3 — Grant the Mail.Read permission

1. In the left sidebar of the app page, click **API permissions**.
2. Click **+ Add a permission** → **Microsoft Graph** → **Delegated
   permissions**.
3. Search for and tick:
   - `Mail.Read`
   - `User.Read` (usually already there)
4. Click **Add permissions**.

That's it for Azure. The first user to sign in will be asked to consent.

## Step 4 — Paste the Client ID into `index.html`

In `index.html`, find this line near the top (around line 1727):

```js
const MS_CLIENT_ID = '';                                  // ← paste your Azure AD app's Application (client) ID here
```

Replace the empty string with the Client ID from Step 2:

```js
const MS_CLIENT_ID = '12345678-1234-1234-1234-123456789abc';
```

Save, commit, push.

## Step 5 — Sign in

1. Open the tracker (`tracker.crymbleandsons.com`)
2. Hard-refresh (Ctrl+Shift+R)
3. Scroll to the **📧 Outlook Inbox** panel under Recent Activity
4. Click **Connect Outlook** — a Microsoft popup opens
5. Sign in, accept the permission prompt
6. The first 15 emails from your inbox appear in the panel
7. Open any order whose customer has an email address — you'll see a
   **📨 Customer Emails** card listing all emails to/from that address

The sign-in is remembered in your browser's local storage. Each user signs
in with their own Microsoft account, so each user sees their own mailbox
and the activity log gets the right name on view events.

## Adding extra users / tenants

The default registration accepts both work/school and personal Microsoft
accounts. If you want to restrict to a single tenant (more secure), change
`MS_TENANT` in `index.html` from `'common'` to your tenant ID (visible on
the Azure app overview page next to **Directory (tenant) ID**).

## Troubleshooting

- **Popup blocked** — allow popups for `tracker.crymbleandsons.com` in your
  browser settings.
- **"Need admin approval" message** — your Microsoft 365 tenant has
  admin-consent-only set up. An IT admin needs to grant the Mail.Read
  permission once for the org via the API permissions page (the
  "Grant admin consent for …" button).
- **Inbox loads but customer emails panel says "0 found"** — Graph search
  matches the address inside email subject/body/recipients/sender. If the
  customer's address differs slightly from what's stored on the order,
  search won't match. Update the email field on the order.
- **403 / 401 errors** — the access token expired and silent renewal
  failed. Click **Sign out** then **Connect Outlook** again.

## Privacy / Permissions

This integration only reads. The app cannot send, delete, or modify any
email. Microsoft enforces this via the `Mail.Read` scope; the app cannot
ask for more without re-registering.
