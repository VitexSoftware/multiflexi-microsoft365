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

namespace MultiFlexi\Ui;

/**
 * Test stub for CredentialFormHelperPrototype (lives in multiflexi-web).
 */
abstract class CredentialFormHelperPrototype
{
    protected mixed $credential = null;

    public function addItem(mixed $item): void
    {
    }

    public function addStatusMessage(mixed $message, string $type = 'info'): void
    {
    }

    public function finalize(): void
    {
    }
}
