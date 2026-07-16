<?php

namespace Tualo\Office\MSGraphVFS;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\MSGraphVFS\StreamWrapper\SharePointStreamWrapper;
use Tualo\Office\MSGraph\API;

class VFS
{
    private static $config = [];

    private static function getConfig(): array
    {
        if (empty(self::$config)) {
            self::$config = [
                'driveId' => App::configuration('msgraph-vfs', 'driveId', ''),
                'siteId' => App::configuration('msgraph-vfs', 'siteId', ''),
            ];
        }
        return self::$config;
    }

    private static function setConfig(string $key, string $value): void
    {
        $config = self::getConfig();
        $config[$key] = $value;
        self::$config = $config;
    }

    private static function normalizeSiteIdentifier(string $siteIdentifier): string
    {
        $siteIdentifier = trim($siteIdentifier);
        if ($siteIdentifier === '') {
            return $siteIdentifier;
        }

        if (str_starts_with($siteIdentifier, 'http://') || str_starts_with($siteIdentifier, 'https://')) {
            $parts = parse_url($siteIdentifier);
            if (is_array($parts) && isset($parts['host'])) {
                $path = isset($parts['path']) ? trim($parts['path'], '/') : '';
                return $path === '' ? $parts['host'] . ':' : $parts['host'] . ':/' . $path . ':';
            }
        }

        return $siteIdentifier;
    }

    private static function getScheme(): string
    {
        $scheme = (string) App::configuration('msgraph-vfs', 'scheme', 'msgraph-vfs');
        return trim($scheme) === '' ? 'msgraph-vfs' : trim($scheme);
    }

    public static function registerVFS(string $siteURL, string $scheme = ''): void
    {
        $scheme = $scheme === '' ? self::getScheme() : $scheme;

        SharePointStreamWrapper::configure(self::resolveSharePointIds($siteURL));

        if (in_array($scheme, stream_get_wrappers(), true)) {
            @stream_wrapper_unregister($scheme);
        }

        if (!stream_wrapper_register($scheme, SharePointStreamWrapper::class)) {
            throw new \RuntimeException(sprintf('Unable to register MS Graph VFS scheme "%s".', $scheme));
        }
    }

    public static function getDriveId(): string
    {
        $configuredDriveId = trim((string) App::configuration('msgraph-vfs', 'driveId', ''));
        if ($configuredDriveId !== '') {
            return $configuredDriveId;
        }

        $siteId = self::normalizeSiteIdentifier((string) App::configuration('msgraph-vfs', 'siteId', ''));
        $graphClient = API::GraphClient();

        if ($siteId !== '') {
            $drive = $graphClient->sites()->bySiteId($siteId)->drive()->get()->wait();
            if ($drive !== null && method_exists($drive, 'getId') && $drive->getId() !== null) {
                return $drive->getId();
            }
        }

        if (API::has('clientSecret')) {
            throw new \RuntimeException('In app-only mode /me is not available. Configure msgraph-vfs.driveId or msgraph-vfs.siteId.');
        }

        $drive = $graphClient->me()->drive()->get()->wait();
        if ($drive !== null && method_exists($drive, 'getId') && $drive->getId() !== null) {
            return $drive->getId();
        }

        throw new \RuntimeException('No MS Graph drive configured. Set msgraph-vfs.driveId or msgraph-vfs.siteId.');
    }

    /**
     * Resolve both the Site ID and Drive ID for a SharePoint site URL.
     *
     * @return array{siteId:?string, driveId:?string}
     */
    public static function resolveSharePointIds(string $siteURL): array
    {
        $siteIdentifier = self::normalizeSiteIdentifier($siteURL);
        if ($siteIdentifier === '') {
            throw new \RuntimeException('A SharePoint site URL is required.');
        }

        $graphClient = API::GraphClient();
        $site = $graphClient->sites()->bySiteId($siteIdentifier)->get()->wait();

        if ($site === null || !method_exists($site, 'getId') || $site->getId() === null) {
            throw new \RuntimeException(sprintf('Unable to resolve site ID for "%s".', $siteURL));
        }

        $siteId = $site->getId();
        $driveId = null;

        try {
            $drive = $graphClient->sites()->bySiteId($siteIdentifier)->drive()->get()->wait();
            if ($drive !== null && method_exists($drive, 'getId') && $drive->getId() !== null) {
                $driveId = $drive->getId();
            }
        } catch (\Throwable $throwable) {
        }

        return [
            'siteId' => $siteId,
            'driveId' => $driveId,
        ];
    }

    public static function getSiteId(?string $siteURL = null): ?string
    {
        $configuredSiteId = trim((string) App::configuration('msgraph-vfs', 'siteId', ''));
        if ($configuredSiteId !== '') {
            return $configuredSiteId;
        }

        $configuredDriveId = trim((string) App::configuration('msgraph-vfs', 'driveId', ''));
        $graphClient = API::GraphClient();

        if ($siteURL !== null && $siteURL !== '') {
            return self::resolveSharePointIds($siteURL)['siteId'];
        }

        if ($configuredDriveId !== '') {
            $drive = $graphClient->drives()->byDriveId($configuredDriveId)->get()->wait();
            if ($drive !== null && method_exists($drive, 'getSharePointIds') && $drive->getSharePointIds() !== null) {
                return $drive->getSharePointIds()->getSiteId();
            }

            return null;
        }

        if (API::has('clientSecret')) {
            return null;
        }

        try {
            $driveId = self::getDriveId();
            $drive = $graphClient->drives()->byDriveId($driveId)->get()->wait();
            if ($drive !== null && method_exists($drive, 'getSharePointIds') && $drive->getSharePointIds() !== null) {
                return $drive->getSharePointIds()->getSiteId();
            }
        } catch (\Throwable $throwable) {
        }

        return null;
    }
}
