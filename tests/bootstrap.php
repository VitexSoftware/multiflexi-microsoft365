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

require_once __DIR__.'/../vendor/autoload.php';

// gettext may be unavailable in the test environment.
if (!\function_exists('_')) {
    function _($text)
    {
        return $text;
    }
}

// Stub for CredentialFormHelperPrototype — lives in multiflexi-web, not required here.
if (!class_exists(\MultiFlexi\Ui\CredentialFormHelperPrototype::class)) {
    require_once __DIR__.'/stubs/MultiFlexi/Ui/CredentialFormHelperPrototype.php';
}
