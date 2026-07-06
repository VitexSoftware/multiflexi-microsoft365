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

## Two-phase connection check

A valid token does **not** guarantee data access: legacy Azure ACS app-only auth can return a
syntactically valid token (HTTP 200) that SharePoint then rejects (HTTP 401 `invalid_request`).
The availability check therefore:

1. Validates required fields are present (offline).
2. **Phase 1** — acquires a token, preferring the modern Entra ID v2 `client_credentials`
   flow (`login.microsoftonline.com/{tenant}/oauth2/v2.0/token`, scope
   `https://{tenant}.sharepoint.com/.default`), falling back to legacy ACS.
3. **Phase 2** — performs a real REST call (`GET /_api/web/title`) with that token.
4. Classifies the outcome as `Available`, `Unavailable` (transient — retry) or
   `Misconfigured` (permanent — needs user action), and reports which phase failed and which
   auth flow worked. The client secret is never logged.

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
