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

        // 2) Phase two — real REST call through php-spo, exactly like the uploader.
        $credentials = $this->credentials();
        $ctx = (new \Office365\SharePoint\ClientContext($this->siteUrl()))->withCredentials($credentials);

        try {
            self::withSharePointRetry($ctx, $credentials, static function ($ctx): void {
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
                ['site' => $this->siteUrl(), 'flow' => $flow],
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
                ['site' => $this->siteUrl(), 'flow' => $flow, 'rest_http' => $restHttp],
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
     * php-spo credential object (user credential preferred when provided).
     *
     * @return \Office365\Runtime\Auth\ClientCredential|\Office365\Runtime\Auth\UserCredentials
     */
    public function credentials()
    {
        if ($this->hasUserCredential()) {
            return new \Office365\Runtime\Auth\UserCredentials($this->username, $this->password);
        }

        return new \Office365\Runtime\Auth\ClientCredential($this->clientId, $this->clientSecret);
    }

    /**
     * Build an authenticated php-spo ClientContext for real data operations.
     */
    public function sharePointContext(): \Office365\SharePoint\ClientContext
    {
        return (new \Office365\SharePoint\ClientContext($this->siteUrl()))->withCredentials($this->credentials());
    }

    /**
     * Run a SharePoint ClientContext operation, retrying once if a transient
     * Microsoft ACS failure leaves the SDK's cached auth token broken.
     *
     * \Office365\Runtime\Auth\ACSTokenProvider hands the decoded response body
     * straight to AuthenticationContext as the access token without checking the
     * HTTP status; a transient ACS hiccup therefore caches an error body as the
     * "token" and every later request on the same context keeps resending it.
     * Calling ClientContext::withCredentials() again installs a fresh
     * AuthenticationContext (token reset to null), forcing a brand new token
     * request on the next call, while $ctx and any object built from it keep
     * working unchanged.
     *
     * @param \Office365\SharePoint\ClientContext                                              $ctx         reset in place on retry
     * @param \Office365\Runtime\Auth\ClientCredential|\Office365\Runtime\Auth\UserCredentials $credentials reinstalled on $ctx before retrying
     * @param callable                                                                         $operation   receives $ctx, performs the call(s) and returns the result
     *
     * @return mixed whatever $operation returned
     */
    public static function withSharePointRetry(
        \Office365\SharePoint\ClientContext $ctx,
        $credentials,
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
                $ctx->withCredentials($credentials);
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
            return $this->tokenProbe = ['ok' => true, 'flow' => 'v2', 'http' => $v2['http'], 'error' => ''];
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
