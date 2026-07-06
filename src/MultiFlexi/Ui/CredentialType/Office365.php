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

namespace MultiFlexi\Ui\CredentialType;

use MultiFlexi\Office365\ProbeResult;
use MultiFlexi\Office365\SharePointConnector;

/**
 * Office 365 / SharePoint credential form helper.
 *
 * Runs the two-phase connection check and renders the outcome (which auth flow
 * worked, per-phase HTTP status and actionable hints). The client secret and
 * password are never rendered.
 *
 * @author Vitex <info@vitexsoftware.cz>
 */
class Office365 extends \MultiFlexi\Ui\CredentialFormHelperPrototype
{
    public function finalize(): void
    {
        $value = function (string $code): string {
            $field = $this->credential->getFields()->getFieldByCode($code);

            return $field ? (string) $field->getValue() : '';
        };

        $tenant = $value('OFFICE365_TENANT');
        $site = $value('OFFICE365_SITE');

        if ($tenant === '' || $site === '') {
            $this->addItem(new \Ease\TWB4\Alert('danger', _('OFFICE365_TENANT and OFFICE365_SITE are required')));

            parent::finalize();

            return;
        }

        $connector = new SharePointConnector(
            $tenant,
            $site,
            $value('OFFICE365_CLIENTID'),
            $value('OFFICE365_CLSECRET'),
            $value('OFFICE365_USERNAME'),
            $value('OFFICE365_PASSWORD'),
            $value('OFFICE365_PATH'),
        );

        $result = $connector->verify();

        $alertType = match ($result->classification) {
            ProbeResult::AVAILABLE => 'success',
            ProbeResult::TRANSIENT => 'warning',
            default => 'danger',
        };

        $panel = new \Ease\TWB4\Panel(_('Microsoft 365 / SharePoint connection'), $alertType);
        $panel->addItem(new \Ease\TWB4\Alert($alertType, $result->message));

        $details = new \Ease\Html\DlTag(null, ['class' => 'row']);
        $rows = [
            _('Site URL') => $connector->siteUrl(),
            _('Auth flow') => strtoupper($result->flow),
            _('Token phase HTTP') => (string) $result->tokenHttp,
            _('REST phase HTTP') => (string) $result->restHttp,
            _('Result') => $result->classification,
            _('Failing phase') => $result->phase,
        ];

        foreach ($rows as $label => $item) {
            $details->addItem(new \Ease\Html\DtTag($label, ['class' => 'col-sm-4']));
            $details->addItem(new \Ease\Html\DdTag($item === '' ? '—' : $item, ['class' => 'col-sm-8']));
        }

        $panel->addItem($details);
        $this->addItem($panel);

        parent::finalize();
    }
}
