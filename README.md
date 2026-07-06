# multiflexi-microsoft365

![MultiFlexi Microsoft 365](multiflexi-microsoft365.svg)

Microsoft 365 / SharePoint integration support for [MultiFlexi](https://multiflexi.eu).

## Description

This package provides the Office 365 credential prototype for MultiFlexi, split into two Debian packages:

| Package | Enhances | Provides |
|---------|----------|----------|
| `multiflexi-microsoft365` | `php-vitexsoftware-multiflexi-core` | Credential prototype with tenant, site, client id/secret and user/password fields plus a live two-phase SharePoint availability check |
| `multiflexi-microsoft365-ui` | `multiflexi-web` | Connection dialog rendering the two-phase probe result, the working auth flow and actionable diagnostics |

SharePoint communication is built on the [`vgrem/php-spo`](https://github.com/vgrem/phpSPO) library (Debian package `php-vgrem-php-spo`).

## Credential Fields

- **OFFICE365_TENANT** — tenant short name (e.g. `contoso` for `contoso.sharepoint.com`)
- **OFFICE365_SITE** — SharePoint site name (`/sites/<SITE>`)
- **OFFICE365_CLIENTID** — application (client) id
- **OFFICE365_CLSECRET** — client secret value
- **OFFICE365_USERNAME** — user principal name (optional, user credential flow)
- **OFFICE365_PASSWORD** — user password (optional, user credential flow)
- **OFFICE365_PATH** — server-relative folder path (optional)

## Authentication

Microsoft fully retired SharePoint "App-Only via Azure ACS" for all tenants on 2026-04-02, with
no extension possible ([retirement announcement](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/retirement-announcement-for-azure-acs)).
Confirmed empirically: the ACS token endpoint still issues a syntactically valid token
(HTTP 200), but SharePoint Online now rejects it on the real REST call
(HTTP 401 `invalid_request`) regardless of credential correctness. Real data operations
(`SharePointConnector::sharePointContext()`) therefore authenticate client-id/secret
credentials via the modern **Entra ID v2 `client_credentials`** flow
(`login.microsoftonline.com/{tenant}/oauth2/v2.0/token`, scope
`https://{tenant}.sharepoint.com/.default`) through `EntraIdAppOnlyAuthenticationContext`,
not ACS. The `OFFICE365_USERNAME`/`OFFICE365_PASSWORD` user-credential flow is unaffected.

The app registration behind `OFFICE365_CLIENTID`/`OFFICE365_CLSECRET` additionally needs the
Microsoft Graph **`Sites.Selected`** application permission, admin-consented, and an explicit
grant on the target site (`POST /sites/{siteId}/permissions`) — a valid token alone is not
sufficient (SharePoint responds `"Unsupported app only token."` otherwise).

## Two-phase connection check

A valid token does **not** guarantee data access. The availability check therefore:

1. Validates required fields are present (offline).
2. **Phase 1** — acquires a token, preferring the modern Entra ID v2 `client_credentials`
   flow, falling back to legacy ACS *for diagnostic comparison only* (real operations never
   use ACS).
3. **Phase 2** — performs a real REST call (`GET /_api/web/title`) through the same
   `sharePointContext()` real operations use, so the check reflects reality.
4. Classifies the outcome as `Available`, `Unavailable` (transient — retry) or
   `Misconfigured` (permanent — needs user action), and reports which phase failed and which
   auth flow worked, plus (best-effort) the access token's and client secret's expiry. The
   client secret is never logged.

### Expiry awareness

- **Access token expiry** is always known once a v2 token is acquired and is shown as
  `token_expires_at` / "Access token expires".
- **Client secret expiry** is looked up via Microsoft Graph
  (`GET /applications?$filter=appId eq '...'`), but this requires the app to additionally hold
  the Graph **`Application.Read.All`** application permission — separate from whatever lets it
  talk to SharePoint. Most SharePoint-only app registrations won't have it, so this is
  best-effort: shown as "Client secret expires (latest on record)" when available, or
  "unknown" with the reason otherwise (never treated as an error).

## Installation

### From Debian packages

```bash
apt install multiflexi-microsoft365 multiflexi-microsoft365-ui
```

### From source (development)

```bash
composer install
make phpunit
make cs
```

## Building Debian Packages

```bash
make deb
```

This produces `multiflexi-microsoft365_*.deb` and `multiflexi-microsoft365-ui_*.deb` in the parent directory.

## License

MIT — see [debian/copyright](debian/copyright) for details.

## MultiFlexi

[![MultiFlexi](https://github.com/VitexSoftware/MultiFlexi/blob/main/doc/multiflexi-app.svg)](https://www.multiflexi.eu/)
