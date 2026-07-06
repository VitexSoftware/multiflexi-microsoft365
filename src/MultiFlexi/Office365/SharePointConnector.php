<?php

declare(strict_types=1);

/**
 * This file is part of the MultiFlexi package
 *
 * https://multiflexi.eu/
 *
 * (c) Vítězslav Dvořák <http://vitexsoftware.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace MultiFlexi\Office365;

/**
 * Office365/SharePoint connection helper built on the vgrem/php-spo library.
 *
 * Provides a trustworthy, two-phase connection check plus the resilient data-op
 * primitives (context builder, retry wrapper, exception describer) used the same
 * way pohoda-raiffeisenbank's SharePoint uploader uses them.
 *
 * Why two phases: the legacy Azure ACS app-only endpoint can return a
 * syntactically valid token (HTTP 200) that SharePoint Online then rejects
 * (HTTP 401 invalid_request) at the tenant level. Checking only "did I get a
 * token" gives a false positive, so verify() always follows a token with a real
 * REST call and reports which phase failed.
 *
 * The client secret is used to authenticate but is never included in any
 * returned diagnostic string.
 *
 * @author Vitex <info@vitexsoftware.cz>
 *
 * @no-named-arguments
 */
class SharePointConnector
{
    /**
     * Memoized token probe result — one independent probe per instance is enough.
     *
     * @var null|array<string, mixed>
     */
    private ?array $tokenProbe = null;

    /**
     * Memoized client secret expiry lookup result — one attempt per instance is enough.
     *
     * @var null|array<string, mixed>
     */
    private ?array $secretExpiryProbe = null;

    public function __construct(
        private readonly string $tenant,
        private readonly string $site,
        private readonly string $clientId = '',
        private readonly string $clientSecret = '',
        private readonly string $username = '',
        private readonly string $password = '',
        private readonly string $path = '',
    ) {
    }

    /**
     * SharePoint host derived from the tenant short name.
     */
    public function spHost(): string
    {
        return $this->tenant.'.sharepoint.com';
    }

    /**
     * Absolute site URL the checks and data operations act against.
     */
    public function siteUrl(): string
    {
        return 'https://'.$this->spHost().'/sites/'.$this->site;
    }

    /**
     * Server-relative folder path for data operations (may be empty).
     */
    public function path(): string
    {
        return $this->path;
    }

    public function hasClientCredential(): bool
    {
        return $this->clientId !== '' && $this->clientSecret !== '';
    }

    public function hasUserCredential(): bool
    {
        return $this->username !== '' && $this->password !== '';
    }

    /**
     * List missing required configuration fields (offline validation).
     *
     * @return array<int, string>
     */
    public function missingRequiredFields(): array
    {
        $missing = [];

        if ($this->tenant === '') {
            $missing[] = 'OFFICE365_TENANT';
        }

        if ($this->site === '') {
            $missing[] = 'OFFICE365_SITE';
        }

        if (!$this->hasClientCredential() && !$this->hasUserCredential()) {
            $missing[] = 'OFFICE365_CLIENTID+OFFICE365_CLSECRET or OFFICE365_USERNAME+OFFICE365_PASSWORD';
        }

        return $missing;
    }

    /**
     * Map an HTTP status / curl errno to a coarse classification.
     *
     * A retry has already been attempted by the caller when relevant, so a 401/403
     * seen here is treated as permanent (SharePoint rejects the token), while 5xx,
     * 429, connection errors and "no response" are transient.
     */
    public static function classify(int $httpCode, int $curlErrno): string
    {
        if ($curlErrno !== 0) {
            return ProbeResult::TRANSIENT;
        }

        if ($httpCode >= 200 && $httpCode < 300) {
            return ProbeResult::AVAILABLE;
        }

        if ($httpCode === 0 || $httpCode === 429 || $httpCode >= 500) {
            return ProbeResult::TRANSIENT;
        }

        return ProbeResult::PERMANENT;
    }

    /**
     * Two-phase connection check: acquire a token, then make a real REST call.
     */
    public function verify(): ProbeResult
    {
        // 0) Offline validation
        $missing = $this->missingRequiredFields();

        if ($missing !== []) {
            return new ProbeResult(
                false,
                ProbeResult::PERMANENT,
                'offline',
                'none',
                0,
                0,
                sprintf(_('Required fields not set: %s'), implode(', ', $missing)),
                ['missing' => implode(',', $missing)],
            );
        }

        // 1) Phase one — token acquisition (client-credential flows only; the user
        // flow is validated directly by the phase-two REST call).
        $flow = 'user';
        $tokenHttp = 0;

        if ($this->hasClientCredential()) {
            $probe = $this->probeToken();
            $tokenHttp = (int) $probe['http'];

            if ($probe['ok']) {
                $flow = (string) $probe['flow'];
            } elseif (!$this->hasUserCredential()) {
                return new ProbeResult(
                    false,
                    self::classify($tokenHttp, 0),
                    'token',
                    'none',
                    $tokenHttp,
                    0,
                    sprintf(
                        _('Token acquisition failed (HTTP %d): %s — check client id/secret/tenant (the secret may be expired)'),
                        $tokenHttp,
                        (string) $probe['error'],
                    ),
                    ['token_error' => (string) $probe['error']],
                );
            }
        }

        // 2) Phase two — real REST call through php-spo, using the same
        // context real data operations use (sharePointContext()), so the
        // check reflects reality instead of a separate, possibly different,
        // auth path.
        $ctx = $this->sharePointContext();
        $resetAuth = $this->resetAuthCallback($ctx);

        try {
            self::withSharePointRetry($ctx, $resetAuth, static function ($ctx): void {
                $web = $ctx->getWeb();
                $ctx->load($web);
                $ctx->executeQuery();
            });

            return new ProbeResult(
                true,
                ProbeResult::AVAILABLE,
                'rest',
                $flow,
                $tokenHttp,
                200,
                sprintf(_('SharePoint reachable at %s (auth flow: %s)'), $this->siteUrl(), $flow),
                $this->withExpiryDetails(['site' => $this->siteUrl(), 'flow' => $flow]),
            );
        } catch (\Office365\Runtime\Http\RequestException $exc) {
            $restHttp = (int) $exc->getCode();

            return new ProbeResult(
                false,
                self::classify($restHttp, 0),
                'rest',
                $flow,
                $tokenHttp,
                $restHttp,
                $this->describeRestFailure($exc),
                $this->withExpiryDetails(['site' => $this->siteUrl(), 'flow' => $flow, 'rest_http' => $restHttp]),
            );
        } catch (\Throwable $exc) {
            return new ProbeResult(
                false,
                ProbeResult::TRANSIENT,
                'rest',
                $flow,
                $tokenHttp,
                0,
                sprintf(_('SharePoint request failed: %s'), $exc->getMessage()),
                ['site' => $this->siteUrl(), 'flow' => $flow],
            );
        }
    }

    /**
     * php-spo credential object for the legacy user-credential flow only.
     * Not used for the client-id/secret case - see sharePointContext().
     */
    public function credentials(): \Office365\Runtime\Auth\UserCredentials
    {
        return new \Office365\Runtime\Auth\UserCredentials($this->username, $this->password);
    }

    /**
     * Build an authenticated php-spo ClientContext for real data operations.
     *
     * Client-id/secret credentials authenticate via the modern Entra ID v2
     * app-only client_credentials flow (EntraIdAppOnlyAuthenticationContext),
     * not the legacy Azure ACS ClientCredential/withCredentials() SDK path -
     * Microsoft fully retired ACS for all tenants on 2026-04-02, so that path
     * no longer works regardless of credential correctness (confirmed: ACS
     * still issues a syntactically valid token, but SharePoint Online rejects
     * it on the real REST call). The user-credential flow is unaffected and
     * unchanged.
     */
    public function sharePointContext(): \Office365\SharePoint\ClientContext
    {
        if ($this->hasUserCredential()) {
            return (new \Office365\SharePoint\ClientContext($this->siteUrl()))->withCredentials($this->credentials());
        }

        $authCtx = new EntraIdAppOnlyAuthenticationContext(
            $this->tenant,
            $this->clientId,
            $this->clientSecret,
            'https://'.$this->spHost().'/.default',
        );

        return new \Office365\SharePoint\ClientContext($this->siteUrl(), $authCtx);
    }

    /**
     * Callback that discards whatever cached auth state caused a failure on
     * $ctx, forcing a fresh attempt on the next request - for use with
     * withSharePointRetry(). Matches how $ctx was built by sharePointContext().
     */
    public function resetAuthCallback(\Office365\SharePoint\ClientContext $ctx): \Closure
    {
        if ($this->hasUserCredential()) {
            $credentials = $this->credentials();

            return static function () use ($ctx, $credentials): void {
                $ctx->withCredentials($credentials);
            };
        }

        return static function () use ($ctx): void {
            $ctx->getAuthenticationContext()->forceRefresh();
        };
    }

    /**
     * Run a SharePoint ClientContext operation, retrying once on a transient
     * auth failure.
     *
     * $resetAuth is called before the retry to discard whatever cached
     * token/credential state caused the failure, forcing a fresh one on the
     * next attempt - see resetAuthCallback(). This matters because, at least
     * for the legacy ACS flow this replaces, a plain resend of the same
     * request does *not* help: \Office365\Runtime\Auth\ACSTokenProvider hands
     * the decoded response body straight to AuthenticationContext as the
     * access token without checking the HTTP status, so a transient hiccup
     * caches a broken "token" that a bare retry would just resubmit unchanged.
     *
     * @param \Office365\SharePoint\ClientContext $ctx          context the operation acts through
     * @param callable                            $resetAuth    called before each retry to force a fresh auth attempt - see resetAuthCallback()
     * @param callable                            $operation    receives $ctx, performs the call(s) and returns the result
     * @param int                                  $maxAttempts  total attempts including the first, before giving up
     * @param int                                  $delaySeconds seconds to wait before a retry
     *
     * @return mixed whatever $operation returned
     */
    public static function withSharePointRetry(
        \Office365\SharePoint\ClientContext $ctx,
        callable $resetAuth,
        callable $operation,
        int $maxAttempts = 2,
        int $delaySeconds = 2,
    ) {
        for ($attempt = 1; ; ++$attempt) {
            try {
                return $operation($ctx);
            } catch (\Office365\Runtime\Http\RequestException $exc) {
                if ($attempt >= $maxAttempts) {
                    throw $exc;
                }

                sleep($delaySeconds);
                $resetAuth();
            }
        }
    }

    /**
     * Build a diagnostic message for a failed SharePoint request.
     *
     * Surfaces the HTTP status and Microsoft's raw response body (which usually
     * carries error_description) plus the independent auth-probe verdict, so the
     * log line alone is enough to diagnose the failure without production access.
     *
     * @param string $context short description of what was attempted
     */
    public function describeRequestException(\Exception $exc, string $context): string
    {
        $detail = $exc->getMessage();
        $source = 'error';

        if ($exc instanceof \Office365\Runtime\Http\RequestException) {
            $source = 'Office365/SharePoint API error';

            if ($exc->getResponseBody()) {
                $detail .= ' | response: '.$exc->getResponseBody();
            }

            $detail .= ' | '.$this->probeSummary();
        }

        return sprintf('%s: %s (HTTP %d): %s', $context, $source, $exc->getCode(), $detail);
    }

    /**
     * Independent, secret-free summary of whether the credentials can obtain a token.
     *
     * Distinguishes "credentials/tenant are broken" from "auth is fine, the REST
     * call itself failed", which is the whole point of the two-phase check.
     */
    public function probeSummary(): string
    {
        if (!$this->hasClientCredential()) {
            return _('auth probe: user credential flow (no app-only token probe performed)');
        }

        $probe = $this->probeToken();

        if ($probe['ok']) {
            return sprintf(
                _('auth probe: independent %s token request succeeded (HTTP %d) — credentials/tenant are valid; the failure is specific to the SharePoint REST call, not auth'),
                strtoupper((string) $probe['flow']),
                (int) $probe['http'],
            );
        }

        return sprintf(
            _('auth probe: independent token request FAILED (HTTP %d): %s'),
            (int) $probe['http'],
            (string) $probe['error'],
        );
    }

    /**
     * Merge token/secret expiry information (when known) into a details array,
     * for the client-credential flow. Best-effort: secret expiry requires the
     * app to additionally hold the Graph "Application.Read.All" application
     * permission, which most SharePoint-only app registrations will not have
     * - that is reported as "unknown", not treated as an error.
     *
     * @param array<string, scalar> $details
     *
     * @return array<string, scalar>
     */
    private function withExpiryDetails(array $details): array
    {
        if (!$this->hasClientCredential()) {
            return $details;
        }

        $tokenProbe = $this->probeToken();

        if ($tokenProbe['ok'] && ($tokenProbe['flow'] ?? '') === 'v2' && !empty($tokenProbe['expires_at'])) {
            $details['token_expires_at'] = date('c', (int) $tokenProbe['expires_at']);
        }

        $secretProbe = $this->probeSecretExpiry();

        if ($secretProbe['ok']) {
            $details['secret_latest_expiry'] = date('c', (int) $secretProbe['expires_at']);
        } else {
            $details['secret_expiry_unknown'] = (string) $secretProbe['error'];
        }

        return $details;
    }

    /**
     * Best-effort lookup of the Entra ID app registration's client secret
     * expiry via Microsoft Graph (GET /applications?$filter=appId eq '...').
     *
     * Requires a Graph token AND the app to hold the Graph
     * "Application.Read.All" application permission - a separate grant from
     * whatever SharePoint/Sites.Selected permission lets it talk to
     * SharePoint itself, so failure here is expected for most app
     * registrations and is reported as "unknown", not as an error state.
     * When an app has multiple secrets on record, the *latest* expiry among
     * them is reported (at least one remains valid until then) - this cannot
     * identify which specific secret is the one actually configured here.
     *
     * @return array<string, mixed> ['ok'=>bool, 'expires_at'=>?int, 'display_name'=>?string, 'error'=>string]
     */
    private function probeSecretExpiry(): array
    {
        if ($this->secretExpiryProbe !== null) {
            return $this->secretExpiryProbe;
        }

        if (!$this->hasClientCredential()) {
            return $this->secretExpiryProbe = ['ok' => false, 'expires_at' => null, 'display_name' => null, 'error' => _('no client credential configured')];
        }

        $authorityTenant = str_contains($this->tenant, '.') ? $this->tenant : $this->tenant.'.onmicrosoft.com';
        [$body, $http] = self::httpPostForm(
            'https://login.microsoftonline.com/'.$authorityTenant.'/oauth2/v2.0/token',
            [
                'grant_type' => 'client_credentials',
                'client_id' => $this->clientId,
                'client_secret' => $this->clientSecret,
                'scope' => 'https://graph.microsoft.com/.default',
            ],
        );
        $decoded = json_decode($body, true);

        if ($http !== 200 || !\is_array($decoded) || !isset($decoded['access_token'])) {
            return $this->secretExpiryProbe = [
                'ok' => false,
                'expires_at' => null,
                'display_name' => null,
                'error' => _('no Microsoft Graph token (app likely lacks a Graph permission grant)'),
            ];
        }

        $ch = curl_init('https://graph.microsoft.com/v1.0/applications?$filter='.rawurlencode("appId eq '{$this->clientId}'").'&$select=displayName,passwordCredentials');
        curl_setopt_array($ch, [
            \CURLOPT_HTTPHEADER => ['Authorization: Bearer '.$decoded['access_token']],
            \CURLOPT_RETURNTRANSFER => true,
            \CURLOPT_TIMEOUT => 10,
        ]);
        $graphBody = curl_exec($ch);
        $graphHttp = (int) curl_getinfo($ch, \CURLINFO_HTTP_CODE);
        curl_close($ch);

        $graphDecoded = \is_string($graphBody) ? json_decode($graphBody, true) : null;
        $app = \is_array($graphDecoded['value'][0] ?? null) ? $graphDecoded['value'][0] : null;

        if ($graphHttp !== 200 || $app === null) {
            return $this->secretExpiryProbe = [
                'ok' => false,
                'expires_at' => null,
                'display_name' => null,
                'error' => sprintf(_('Graph application lookup failed (HTTP %d) - requires the "Application.Read.All" Graph application permission'), $graphHttp),
            ];
        }

        $latestExpiry = null;

        foreach ((array) ($app['passwordCredentials'] ?? []) as $cred) {
            $end = isset($cred['endDateTime']) ? strtotime((string) $cred['endDateTime']) : false;

            if ($end !== false && ($latestExpiry === null || $end > $latestExpiry)) {
                $latestExpiry = $end;
            }
        }

        return $this->secretExpiryProbe = [
            'ok' => $latestExpiry !== null,
            'expires_at' => $latestExpiry,
            'display_name' => (string) ($app['displayName'] ?? ''),
            'error' => $latestExpiry === null ? _('application has no client secrets on record') : '',
        ];
    }

    /**
     * Try to obtain an app-only token, preferring Entra ID v2 and falling back to
     * legacy ACS. Memoized per instance. Never returns the token value itself.
     *
     * @return array<string, mixed>
     */
    private function probeToken(): array
    {
        if ($this->tokenProbe !== null) {
            return $this->tokenProbe;
        }

        $v2 = $this->acquireTokenV2();

        if ($v2['ok']) {
            return $this->tokenProbe = ['ok' => true, 'flow' => 'v2', 'http' => $v2['http'], 'error' => '', 'expires_at' => $v2['expires_at']];
        }

        $acs = $this->acquireTokenAcs();

        if ($acs['ok']) {
            return $this->tokenProbe = ['ok' => true, 'flow' => 'acs', 'http' => $acs['http'], 'error' => ''];
        }

        return $this->tokenProbe = [
            'ok' => false,
            'flow' => 'none',
            'http' => $v2['http'] ?: $acs['http'],
            'error' => sprintf('v2: %s; acs: %s', (string) $v2['error'], (string) $acs['error']),
        ];
    }

    /**
     * Modern Entra ID v2 client_credentials grant (migration-ready path).
     *
     * @return array<string, mixed>
     */
    private function acquireTokenV2(): array
    {
        if (!$this->hasClientCredential()) {
            return ['ok' => false, 'http' => 0, 'errno' => 0, 'error' => 'no client credential'];
        }

        $authorityTenant = str_contains($this->tenant, '.') ? $this->tenant : $this->tenant.'.onmicrosoft.com';
        [$body, $http, $errno, $err] = self::httpPostForm(
            'https://login.microsoftonline.com/'.$authorityTenant.'/oauth2/v2.0/token',
            [
                'grant_type' => 'client_credentials',
                'client_id' => $this->clientId,
                'client_secret' => $this->clientSecret,
                'scope' => 'https://'.$this->spHost().'/.default',
            ],
        );
        $decoded = json_decode($body, true);
        $ok = $http === 200 && \is_array($decoded) && isset($decoded['access_token']);

        return [
            'ok' => $ok,
            'http' => $http,
            'errno' => $errno,
            'error' => $ok ? '' : (\is_array($decoded) && isset($decoded['error']) ? (string) $decoded['error'] : ($err !== '' ? $err : 'HTTP '.$http)),
            'expires_at' => $ok ? time() + (int) ($decoded['expires_in'] ?? 3600) : 0,
        ];
    }

    /**
     * Legacy SharePoint App-Only via Azure ACS (fallback path).
     *
     * @return array<string, mixed>
     */
    private function acquireTokenAcs(): array
    {
        if (!$this->hasClientCredential()) {
            return ['ok' => false, 'http' => 0, 'errno' => 0, 'error' => 'no client credential'];
        }

        $realmInfo = $this->discoverRealm();
        $realm = $realmInfo['realm'];

        if ($realm === null) {
            return ['ok' => false, 'http' => (int) $realmInfo['http'], 'errno' => (int) $realmInfo['errno'], 'error' => 'could not discover ACS realm'];
        }

        $resource = '00000003-0000-0ff1-ce00-000000000000/'.$this->spHost().'@'.$realm;
        [$body, $http, $errno, $err] = self::httpPostForm(
            'https://accounts.accesscontrol.windows.net/'.$realm.'/tokens/OAuth/2',
            [
                'grant_type' => 'client_credentials',
                'client_id' => $this->clientId.'@'.$realm,
                'client_secret' => $this->clientSecret,
                'resource' => $resource,
                'scope' => $resource,
            ],
        );
        $decoded = json_decode($body, true);
        $ok = $http === 200 && \is_array($decoded) && isset($decoded['access_token']);

        return [
            'ok' => $ok,
            'http' => $http,
            'errno' => $errno,
            'error' => $ok ? '' : (\is_array($decoded) && isset($decoded['error']) ? (string) $decoded['error'] : ($err !== '' ? $err : 'HTTP '.$http)),
        ];
    }

    /**
     * Discover the ACS realm from the SharePoint site's WWW-Authenticate header.
     *
     * @return array<string, mixed>
     */
    private function discoverRealm(): array
    {
        $headerLines = [];
        $ch = curl_init($this->siteUrl());
        curl_setopt_array($ch, [
            \CURLOPT_NOBODY => true,
            \CURLOPT_HTTPHEADER => ['Authorization: Bearer'],
            \CURLOPT_HEADERFUNCTION => static function ($ch, $line) use (&$headerLines) {
                $headerLines[] = $line;

                return \strlen($line);
            },
            \CURLOPT_TIMEOUT => 10,
        ]);
        curl_exec($ch);
        $http = (int) curl_getinfo($ch, \CURLINFO_HTTP_CODE);
        $errno = curl_errno($ch);
        curl_close($ch);

        $realm = null;

        foreach ($headerLines as $line) {
            if (str_starts_with(strtolower($line), strtolower('WWW-Authenticate:')) && preg_match('/realm="([^"]+)"/', $line, $m)) {
                $realm = $m[1];

                break;
            }
        }

        return ['realm' => $realm, 'http' => $http, 'errno' => $errno];
    }

    /**
     * Build the phase-two failure message: HTTP status + Microsoft's raw body +
     * the independent auth-probe verdict + actionable hints for 401/403.
     */
    private function describeRestFailure(\Office365\Runtime\Http\RequestException $exc): string
    {
        $http = (int) $exc->getCode();
        $message = sprintf(_('SharePoint rejected the request (HTTP %d)'), $http);

        if ($exc->getResponseBody()) {
            $message .= ' | '.$exc->getResponseBody();
        }

        $message .= ' | '.$this->probeSummary();

        if ($http === 401 || $http === 403) {
            $message .= ' | '._('A token was issued but the REST call was rejected: the tenant likely disabled app-only auth (Get-SPOTenant DisableCustomAppAuthentication), the app-only grant at _layouts/15/appinv.aspx is missing/expired, or the app needs Entra ID v2 with SharePoint Sites.Selected.');
        }

        return $message;
    }

    /**
     * POST an application/x-www-form-urlencoded body and return [body, http, errno, error].
     *
     * @param array<string, string> $fields
     *
     * @return array{0:string, 1:int, 2:int, 3:string}
     */
    private static function httpPostForm(string $url, array $fields): array
    {
        $ch = curl_init($url);
        curl_setopt_array($ch, [
            \CURLOPT_POST => true,
            \CURLOPT_RETURNTRANSFER => true,
            \CURLOPT_HTTPHEADER => ['Content-Type: application/x-www-form-urlencoded'],
            \CURLOPT_POSTFIELDS => http_build_query($fields),
            \CURLOPT_TIMEOUT => 10,
        ]);
        $body = curl_exec($ch);
        $http = (int) curl_getinfo($ch, \CURLINFO_HTTP_CODE);
        $errno = curl_errno($ch);
        $err = curl_error($ch);
        curl_close($ch);

        return [\is_string($body) ? $body : '', $http, $errno, $err];
    }
}
