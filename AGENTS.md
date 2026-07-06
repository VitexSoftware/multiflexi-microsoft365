# AGENTS.md

## Project Overview

This project provides Microsoft 365 / SharePoint integration support for MultiFlexi as a separate Debian-packaged addon. It produces two binary packages from one source:

- **multiflexi-microsoft365** — credential prototype for `php-vitexsoftware-multiflexi-core` (`MultiFlexi\CredentialProtoType\Office365`)
- **multiflexi-microsoft365-ui** — UI form helper for `multiflexi-web` (`MultiFlexi\Ui\CredentialType\Office365`)

## Directory Structure

- `src/MultiFlexi/CredentialProtoType/Office365.php` — core credential prototype class (fields + `checkAvailability()`)
- `src/MultiFlexi/Office365/SharePointConnector.php` — framework-agnostic SharePoint auth/probe/retry helper (uses `vgrem/php-spo`)
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
Parameterized helper (no framework config coupling). `acquireToken()` prefers the Entra ID v2 `client_credentials` grant and falls back to legacy ACS; `verify()` runs the two-phase check; `withSharePointRetry()` resets the SDK's poisoned token cache and retries once; `describeRequestException()` surfaces the HTTP status and Microsoft's raw error body. The client secret is never logged.

### MultiFlexi\Ui\CredentialType\Office365
Extends `\MultiFlexi\Ui\CredentialFormHelperPrototype`.
Renders the two-phase result (probed site URL, working auth flow, per-phase HTTP codes) and actionable hints (tenant `DisableCustomAppAuthentication`, `appinv.aspx` grant, secret expiry).
