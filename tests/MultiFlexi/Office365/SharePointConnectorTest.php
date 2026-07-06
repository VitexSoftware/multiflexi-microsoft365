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

namespace Test\MultiFlexi\Office365;

use MultiFlexi\Office365\ProbeResult;
use MultiFlexi\Office365\SharePointConnector;
use PHPUnit\Framework\TestCase;

class SharePointConnectorTest extends TestCase
{
    public function testClassifyAvailable(): void
    {
        self::assertSame(ProbeResult::AVAILABLE, SharePointConnector::classify(200, 0));
        self::assertSame(ProbeResult::AVAILABLE, SharePointConnector::classify(204, 0));
    }

    public function testClassifyTransient(): void
    {
        self::assertSame(ProbeResult::TRANSIENT, SharePointConnector::classify(0, 0));
        self::assertSame(ProbeResult::TRANSIENT, SharePointConnector::classify(500, 0));
        self::assertSame(ProbeResult::TRANSIENT, SharePointConnector::classify(503, 0));
        self::assertSame(ProbeResult::TRANSIENT, SharePointConnector::classify(429, 0));
        // A curl error is transient regardless of HTTP status.
        self::assertSame(ProbeResult::TRANSIENT, SharePointConnector::classify(200, 7));
    }

    public function testClassifyPermanent(): void
    {
        self::assertSame(ProbeResult::PERMANENT, SharePointConnector::classify(401, 0));
        self::assertSame(ProbeResult::PERMANENT, SharePointConnector::classify(403, 0));
        self::assertSame(ProbeResult::PERMANENT, SharePointConnector::classify(404, 0));
    }

    public function testSiteUrlAndHost(): void
    {
        $connector = new SharePointConnector('contoso', 'MySite');
        self::assertSame('contoso.sharepoint.com', $connector->spHost());
        self::assertSame('https://contoso.sharepoint.com/sites/MySite', $connector->siteUrl());
    }

    public function testMissingRequiredFieldsWhenEmpty(): void
    {
        $connector = new SharePointConnector('', '');
        $missing = $connector->missingRequiredFields();
        self::assertContains('OFFICE365_TENANT', $missing);
        self::assertContains('OFFICE365_SITE', $missing);
        self::assertCount(3, $missing);
    }

    public function testMissingRequiredFieldsSatisfiedWithClientCredential(): void
    {
        $connector = new SharePointConnector('contoso', 'MySite', 'client-id', 'secret');
        self::assertSame([], $connector->missingRequiredFields());
        self::assertTrue($connector->hasClientCredential());
        self::assertFalse($connector->hasUserCredential());
    }

    public function testMissingRequiredFieldsSatisfiedWithUserCredential(): void
    {
        $connector = new SharePointConnector('contoso', 'MySite', '', '', 'user@contoso.onmicrosoft.com', 'pass');
        self::assertSame([], $connector->missingRequiredFields());
        self::assertTrue($connector->hasUserCredential());
        self::assertFalse($connector->hasClientCredential());
    }

    public function testVerifyOfflineIsPermanentWithoutNetwork(): void
    {
        $connector = new SharePointConnector('', '');
        $result = $connector->verify();
        self::assertInstanceOf(ProbeResult::class, $result);
        self::assertFalse($result->ok);
        self::assertSame(ProbeResult::PERMANENT, $result->classification);
        self::assertSame('offline', $result->phase);
        self::assertStringContainsString('OFFICE365_TENANT', $result->message);
    }

    public function testProbeSummaryUserFlowPerformsNoNetworkProbe(): void
    {
        $connector = new SharePointConnector('contoso', 'MySite', '', '', 'user@contoso.onmicrosoft.com', 'pass');
        self::assertStringContainsString('user credential flow', $connector->probeSummary());
    }

    public function testPathAccessor(): void
    {
        $connector = new SharePointConnector('contoso', 'MySite', '', '', '', '', 'Shared Documents/Invoices');
        self::assertSame('Shared Documents/Invoices', $connector->path());
    }
}
