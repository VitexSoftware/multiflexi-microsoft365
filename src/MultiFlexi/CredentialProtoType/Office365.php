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

namespace MultiFlexi\CredentialProtoType;

use MultiFlexi\Office365\ProbeResult;
use MultiFlexi\Office365\SharePointConnector;

/**
 * Office 365 / SharePoint credential prototype.
 *
 * author Vitex <info@vitexsoftware.cz>
 *
 * @no-named-arguments
 */
class Office365 extends \MultiFlexi\CredentialProtoType implements \MultiFlexi\checkableCredentialInterface, \MultiFlexi\credentialTypeInterface
{
    public static string $logo = 'Office365.svg';

    public function __construct()
    {
        parent::__construct();

        $tenantField = new \MultiFlexi\ConfigField('OFFICE365_TENANT', 'string', _('Office 365 Tenant'), _('Tenant short name, e.g. contoso for contoso.sharepoint.com'));
        $tenantField->setHint('contoso')->setRequired(true)->setValue('');

        $siteField = new \MultiFlexi\ConfigField('OFFICE365_SITE', 'string', _('SharePoint Site'), _('SharePoint site name (the SITE part of /sites/SITE)'));
        $siteField->setHint('MySite')->setRequired(true)->setValue('');

        $clientIdField = new \MultiFlexi\ConfigField('OFFICE365_CLIENTID', 'string', _('Office 365 Client ID'), _('Application (client) ID for app-only access'));
        $clientIdField->setHint('00000000-0000-0000-0000-000000000000')->setValue('');

        $clientSecretField = new \MultiFlexi\ConfigField('OFFICE365_CLSECRET', 'password', _('Office 365 Secret Value'), _('Client secret VALUE (Azure Portal → App registrations → Certificates & secrets)'));
        $clientSecretField->setHint('your-secret')->setSecret(true)->setValue('');

        $usernameField = new \MultiFlexi\ConfigField('OFFICE365_USERNAME', 'string', _('Office 365 Username'), _('User principal name (optional user credential flow)'));
        $usernameField->setHint('user@contoso.onmicrosoft.com')->setValue('');

        $passwordField = new \MultiFlexi\ConfigField('OFFICE365_PASSWORD', 'password', _('Office 365 Password'), _('Password for the user credential flow (optional)'));
        $passwordField->setHint('your-password')->setSecret(true)->setValue('');

        $pathField = new \MultiFlexi\ConfigField('OFFICE365_PATH', 'string', _('SharePoint Folder Path'), _('Server-relative folder path for data operations (optional)'));
        $pathField->setHint('Shared Documents/Invoices')->setValue('');

        $this->configFieldsInternal->addField($tenantField);
        $this->configFieldsInternal->addField($siteField);
        $this->configFieldsInternal->addField($clientIdField);
        $this->configFieldsInternal->addField($clientSecretField);
        $this->configFieldsInternal->addField($usernameField);
        $this->configFieldsInternal->addField($passwordField);
        $this->configFieldsInternal->addField($pathField);
    }

    public function load(int $credTypeId)
    {
        $loaded = parent::load($credTypeId);

        // Bridge internal (loaded) credential fields into the provided set.
        foreach ($this->configFieldsInternal->getFields() as $field) {
            $this->configFieldsProvided->addField($field);
        }

        return $loaded;
    }

    #[\Override]
    public function prepareConfigForm(): void
    {
        // No dynamic form logic required.
    }

    public function name(): string
    {
        return _('Office 365');
    }

    public function description(): string
    {
        return _('Credential type for integration with Office 365 / SharePoint');
    }

    #[\Override]
    public function uuid(): string
    {
        return 'd510bdee-de98-47e8-96b2-0b301af7d96b';
    }

    #[\Override]
    public function logo(): string
    {
        return self::$logo;
    }

    #[\Override]
    public function checkAvailability(): \MultiFlexi\CredentialCheckResult
    {
        $f = fn (string $code): string => (string) ($this->configFieldsInternal->getFieldByCode($code)?->getValue() ?? '');

        $connector = new SharePointConnector(
            $f('OFFICE365_TENANT'),
            $f('OFFICE365_SITE'),
            $f('OFFICE365_CLIENTID'),
            $f('OFFICE365_CLSECRET'),
            $f('OFFICE365_USERNAME'),
            $f('OFFICE365_PASSWORD'),
            $f('OFFICE365_PATH'),
        );

        $result = $connector->verify();

        $state = match ($result->classification) {
            ProbeResult::AVAILABLE => \MultiFlexi\CredentialState::Available,
            ProbeResult::TRANSIENT => \MultiFlexi\CredentialState::Unavailable,
            default => \MultiFlexi\CredentialState::Misconfigured,
        };

        $ttl = $state === \MultiFlexi\CredentialState::Unavailable ? 60 : 300;

        $details = array_map(static fn ($value): string => (string) $value, $result->details);

        return new \MultiFlexi\CredentialCheckResult($state, $result->message, time(), $ttl, $details);
    }
}
