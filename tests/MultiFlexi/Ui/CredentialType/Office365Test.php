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

namespace Test\MultiFlexi\Ui\CredentialType;

use MultiFlexi\Ui\CredentialType\Office365;
use PHPUnit\Framework\TestCase;

class Office365Test extends TestCase
{
    public function testClassExists(): void
    {
        self::assertTrue(class_exists(Office365::class));
    }

    public function testExtendsCredentialFormHelperPrototype(): void
    {
        $reflection = new \ReflectionClass(Office365::class);
        self::assertTrue($reflection->isSubclassOf(\MultiFlexi\Ui\CredentialFormHelperPrototype::class));
    }

    public function testHasFinalizeMethod(): void
    {
        self::assertTrue(method_exists(Office365::class, 'finalize'));
    }
}
