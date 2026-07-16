<?php

namespace Tualo\Office\MSGraphVFS;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\MSGraphVFS\StreamWrapper\SharePointStreamWrapper;

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
}
