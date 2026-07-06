# AGENTS.md

## Project Overview

This project provides Microsoft 365 / SharePoint integration support for MultiFlexi as a separate Debian-packaged addon. It produces two binary packages from one source:

- **multiflexi-microsoft365** — credential prototype for `php-vitexsoftware-multiflexi-core` (`MultiFlexi\CredentialProtoType\Office365`)
- **multiflexi-microsoft365-ui** — UI form helper for `multiflexi-web` (`MultiFlexi\Ui\CredentialType\Office365`)

## Directory Structure

- `src/MultiFlexi/CredentialProtoType/Office365.php` — core credential prototype class (fields + `checkAvailability()`)
- `src/MultiFlexi/Office365/SharePointConnector.php` — framework-agnostic SharePoint auth/probe/retry helper (uses `vgrem/php-spo`)
- `src/MultiFlexi/Office365/EntraIdAppOnlyAuthenticationContext.php` — Entra ID v2 app-only `client_credentials` auth, implements `Office365\Runtime\Auth\IAuthenticationContext` directly so it can be passed to `new ClientContext($url, $authCtx)`, replacing the retired Azure ACS flow for real data operations
- `src/MultiFlexi/Office365/ProbeResult.php` — immutable two-phase probe outcome
- `src/MultiFlexi/Ui/CredentialType/Office365.php` — web UI credential form helper
- `src/images/Office365.svg` — logo asset
- `multiflexi/office365.credprototype.json` — reference JSON definition (not installed; sync reads the PHP class)
- `debian/` — Debian packaging
- `tests/` — PHPUnit tests

## Build & Test

```bash
make vendor    # install composer dependencies
make phpunit   # run tests
make cs        # fix coding standards
make deb       # build Debian packages
```

## Coding Standards

- PHP 8.1+ with strict types
- PSR-12 via ergebnis/php-cs-fixer-config
- Run `make cs` before committing

## Debian Packaging

The `debian/control` defines two binary packages with proper dependency chains:
- `multiflexi-microsoft365` depends on `php-vitexsoftware-multiflexi-core (>= 2.9)`, `php-vgrem-php-spo` and `multiflexi-cli (>= 2.2.0)`
- `multiflexi-microsoft365-ui` depends on `multiflexi-microsoft365` and `multiflexi-web|multiflexi-web5`

The `postinst` for `multiflexi-microsoft365` runs `multiflexi-cli credential-prototype:sync` to register the credential prototype. Registration is discovered from the installed PHP class, so the `*.credprototype.json` file is not installed.

## Key Classes

### MultiFlexi\CredentialProtoType\Office365
Extends `\MultiFlexi\CredentialProtoType` and implements `\MultiFlexi\credentialTypeInterface` and `\MultiFlexi\checkableCredentialInterface`.
Defines fields (in `configFieldsInternal`, bridged to `configFieldsProvided` in `load()`): OFFICE365_TENANT, OFFICE365_SITE, OFFICE365_CLIENTID, OFFICE365_CLSECRET, OFFICE365_USERNAME, OFFICE365_PASSWORD, OFFICE365_PATH.
`checkAvailability()` delegates to `SharePointConnector::verify()` for a two-phase (token + real REST call) probe and maps the outcome to `CredentialState`.

### MultiFlexi\Office365\SharePointConnector
Parameterized helper (no framework config coupling).
- `sharePointContext()` builds the authenticated `ClientContext` used for **real** data operations: client-id/secret credentials go through `EntraIdAppOnlyAuthenticationContext` (Entra ID v2 app-only `client_credentials`, scope `https://{tenant}.sharepoint.com/.default`); the legacy `OFFICE365_USERNAME`/`OFFICE365_PASSWORD` user-credential flow is unchanged. Azure ACS (`ClientCredential`/`withCredentials()`) is **not** used here - Microsoft fully retired it for all tenants on 2026-04-02 (confirmed empirically: ACS still issues a syntactically valid token, but SharePoint Online rejects it on the real REST call regardless of credential correctness).
- `probeToken()`/`acquireTokenV2()`/`acquireTokenAcs()` remain diagnostic-only (used by `verify()`/`probeSummary()`) and still try ACS purely for comparison in the report - real operations never do.
- `verify()` runs the two-phase (token + real REST call) probe through `sharePointContext()`, so the check reflects exactly what real operations do; also merges token/secret expiry into `ProbeResult::$details`.
- `withSharePointRetry(ClientContext $ctx, callable $resetAuth, callable $operation, ...)` retries once on a transient auth failure; `$resetAuth` is built by `resetAuthCallback($ctx)` (`EntraIdAppOnlyAuthenticationContext::forceRefresh()` for the v2 case, `ClientContext::withCredentials()` for the user-credential case).
- `describeRequestException()` surfaces the HTTP status and Microsoft's raw error body.
- `probeSecretExpiry()` is a best-effort Microsoft Graph lookup (`GET /applications?$filter=appId eq '...'`) of the app registration's client secret expiry; requires the app to additionally hold the Graph "Application.Read.All" application permission (separate from whatever lets it talk to SharePoint), so failure is expected and reported as "unknown" via `ProbeResult::$details['secret_expiry_unknown']`, not as an error.
- The client secret is never logged.

### MultiFlexi\Ui\CredentialType\Office365
Extends `\MultiFlexi\Ui\CredentialFormHelperPrototype`.
Renders the two-phase result (probed site URL, working auth flow, per-phase HTTP codes), access-token and client-secret expiry when known, and actionable hints (tenant `DisableCustomAppAuthentication`, `appinv.aspx` grant / Graph `Sites.Selected` site grant, secret expiry).
