<?php

namespace Tualo\Office\MSGraphVFS\Routes;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\Basic\Route as BasicRoute;
use Tualo\Office\MSGraphVFS\VFS;

class Test extends \Tualo\Office\Basic\RouteWrapper
{
    public static function register()
    {
        BasicRoute::add('/msgraph-vfs/(?P<urlEncodedSharepointURL>.+)', function ($matches) {
            App::contenttype('application/json');

            try {
                $sharePointUrl = urldecode($matches['urlEncodedSharepointURL']);
                $resolved = VFS::resolveSharePointIds($sharePointUrl);
                App::result('resolved', $resolved);

                VFS::registerVFS($sharePointUrl);

                $scheme = App::configuration('msgraph-vfs', 'scheme', 'msgraph-vfs');
                $entries = [];
                $handle = opendir($scheme . '://');
                if ($handle !== false) {
                    while (($entry = readdir($handle)) !== false) {
                        if ($entry !== '.' && $entry !== '..') {
                            $entries[] = $entry;
                        }
                    }
                    closedir($handle);
                }

                App::result('data', $entries);
                App::result('success', true);
            } catch (\Exception $e) {
                App::result('error', $e->getMessage());
            }
        }, ['get'], true);
    }
}
