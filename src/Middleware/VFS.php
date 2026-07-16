<?php

namespace Tualo\Office\MSGraphVFS\Middleware;

use Tualo\Office\Basic\TualoApplication;
use Tualo\Office\Basic\IMiddleware;
use Tualo\Office\MSGraphVFS\VFS as V;

class VFS implements IMiddleware
{
    public static function register()
    {
        TualoApplication::use('TualoApplication_msgraph_vfs', function () {
            V::registerVFS();
        }, 0, [], true);
    }
}
