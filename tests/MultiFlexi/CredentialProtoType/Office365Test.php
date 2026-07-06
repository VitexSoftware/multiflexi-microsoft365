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

namespace Test\MultiFlexi\CredentialProtoType;

use MultiFlexi\CredentialProtoType\Office365;
use PHPUnit\Framework\TestCase;

class Office365Test extends TestCase
{
    public function testName(): void
    {
        self::assertNotEmpty(self::prototype()->name());
    }

    public function testDescription(): void
    {
        self::assertNotEmpty(self::prototype()->description());
    }

    public function testUuid(): void
    {
        self::assertMatchesRegularExpression(
            '/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/',
            self::prototype()->uuid(),
        );
    }

    public function testLogo(): void
    {
        self::assertSame('Office365.svg', self::prototype()->logo());
    }

    public function testImplementsCheckableInterface(): void
    {
        self::assertInstanceOf(\MultiFlexi\checkableCredentialInterface::class, self::prototype());
    }

    public function testCheckAvailabilityMisconfiguredWhenEmpty(): void
    {
        try {
            $prototype = new Office365();
        } catch (\Throwable $e) {
            self::markTestSkipped('Cannot construct Office365 (Engine/DB unavailable): '.$e->getMessage());
        }

        // No credentials configured → offline validation must report Misconfigured
        // without performing any network request.
        $result = $prototype->checkAvailability();
        self::assertSame(\MultiFlexi\CredentialState::Misconfigured, $result->state);
    }
    /**
     * Metadata methods do not depend on constructor/database state, so build the
     * object without invoking the DB-backed Engine constructor.
     */
    private static function prototype(): Office365
    {
        return (new \ReflectionClass(Office365::class))->newInstanceWithoutConstructor();
    }
}
