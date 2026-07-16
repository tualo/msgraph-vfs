<?php

namespace Tualo\Office\MSGraphVFS;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\MSGraphVFS\StreamWrapper\SharePointStreamWrapper;
use Tualo\Office\MSGraph\API;

class VFS
{
    private static function getScheme(): string
    {
        $scheme = (string) App::configuration('msgraph-vfs', 'scheme', 'msgraph-vfs');
        return trim($scheme) === '' ? 'msgraph-vfs' : trim($scheme);
    }

    public static function registerVFS(): void
    {
        $scheme = self::getScheme();

        SharePointStreamWrapper::configure([
            'driveId' => App::configuration('msgraph-vfs', 'driveId', ''),
            'siteId' => App::configuration('msgraph-vfs', 'siteId', ''),
        ]);

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

        $siteId = trim((string) App::configuration('msgraph-vfs', 'siteId', ''));
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

    public static function getSiteId(): ?string
    {
        $configuredSiteId = trim((string) App::configuration('msgraph-vfs', 'siteId', ''));
        if ($configuredSiteId !== '') {
            return $configuredSiteId;
        }

        $configuredDriveId = trim((string) App::configuration('msgraph-vfs', 'driveId', ''));
        $graphClient = API::GraphClient();

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
