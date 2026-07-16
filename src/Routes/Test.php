<?php

namespace Tualo\Office\MSGraphVFS\Routes;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\Basic\RouteSecurityHelper;
use Tualo\Office\Basic\Route as BasicRoute;
use Tualo\Office\Basic\IRoute;
use Tualo\Office\MSGraphVFS\VFS;

class Test extends \Tualo\Office\Basic\RouteWrapper
{
    public static function register()
    {
        BasicRoute::add('/msgraph-vfs/(?P<file>[\w.\/\-]+.js)', function ($matches) {
            App::contenttype('application/json');
            try {
                // App::result('drive', VFS::getDriveId());
                App::result('site', VFS::getSiteId('https://tualo.sharepoint.com/sites/wvd-tualo'));

                App::result('success', true);
            } catch (\Exception $e) {
                App::result('error', $e->getMessage());
            }
        }, ['get'], true);
    }
}
