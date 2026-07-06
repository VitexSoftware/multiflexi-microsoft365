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
 * Immutable outcome of a two-phase SharePoint connection probe.
 *
 * @author Vitex <info@vitexsoftware.cz>
 */
final class ProbeResult
{
    /**
     * Service reachable and credential valid.
     */
    public const AVAILABLE = 'available';

    /**
     * Transient failure (timeout, 5xx, network, ACS hiccup) — retry later.
     */
    public const TRANSIENT = 'transient';

    /**
     * Permanent failure (bad/expired secret, tenant rejects app-only) — needs user action.
     */
    public const PERMANENT = 'permanent';
    /**
     * @param bool                  $ok             whether SharePoint data access succeeded
     * @param string                $classification one of self::AVAILABLE, self::TRANSIENT, self::PERMANENT
     * @param string                $phase          which phase produced the verdict: offline|token|rest
     * @param string                $flow           auth flow that worked: v2|acs|user|none
     * @param int                   $tokenHttp      HTTP status observed while acquiring the token (0 if not attempted)
     * @param int                   $restHttp       HTTP status of the real REST call (0 if not attempted)
     * @param string                $message        human readable, secret-free summary
     * @param array<string, scalar> $details        extra machine-readable payload for the UI
     */
    public function __construct(
        public readonly bool $ok,
        public readonly string $classification,
        public readonly string $phase,
        public readonly string $flow,
        public readonly int $tokenHttp,
        public readonly int $restHttp,
        public readonly string $message,
        public readonly array $details = [],
    ) {
    }
}
