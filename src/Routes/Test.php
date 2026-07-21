<?php

namespace Tualo\Office\MSGraphVFS\Routes;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\Basic\Route as BasicRoute;
use Tualo\Office\MSGraphVFS\VFS;

class Test extends \Tualo\Office\Basic\RouteWrapper
{
    public static function register()
    {
        BasicRoute::add('/msgraph-vfs(?:/(?P<folderPath>.*))?$', function ($matches) {
            App::contenttype('application/json');

            try {
                // $sharePointUrl = urldecode($matches['urlEncodedSharepointURL']);
                $sharePointUrl = "https://tualo.sharepoint.com/sites/wvd-tualo/";
                $resolved = VFS::resolveSharePointIds($sharePointUrl);
                App::result('resolved', $resolved);

                VFS::registerVFS($sharePointUrl);

                $scheme = App::configuration('msgraph-vfs', 'scheme', 'msgraph-vfs');
                $folderPath = isset($matches['folderPath']) ? trim(urldecode($matches['folderPath']), '/') : '';
                $folderPath = preg_replace('#/Forms/AllItems\.aspx$#i', '', $folderPath) ?? $folderPath;
                $folderPath = preg_replace('#/Forms$#i', '', $folderPath) ?? $folderPath;
                $folderSegments = $folderPath === '' ? [] : explode('/', $folderPath);
                if (count($folderSegments) > 0) {
                    array_shift($folderSegments);
                }
                $vfsPath = implode('/', $folderSegments);
                $directoryUrl = $scheme . '://' . ($vfsPath !== '' ? $vfsPath : '');

                if ($vfsPath !== '' && !is_dir($directoryUrl)) {
                    mkdir($directoryUrl, 0777, true);
                }

                $targetFile = $directoryUrl . '/test.txt';
                $bytesWritten = file_put_contents($targetFile, 'Inhalt');
                if ($bytesWritten === false) {
                    throw new \RuntimeException(sprintf('Writing %s failed.', $targetFile));
                }

                clearstatcache(true, $targetFile);
                $fileExists = file_exists($targetFile);
                if (!$fileExists) {
                    throw new \RuntimeException(sprintf('File %s was not created.', $targetFile));
                }

                $entries = [];
                $handle = opendir($directoryUrl);
                if ($handle !== false) {
                    while (($entry = readdir($handle)) !== false) {
                        if ($entry !== '.' && $entry !== '..') {
                            $entries[] = $entry;
                        }
                    }
                    closedir($handle);
                }

                App::result('writtenBytes', $bytesWritten);
                App::result('fileExists', $fileExists);
                App::result('data', $entries);
                App::result('success', true);
            } catch (\Exception $e) {
                App::result('error', $e->getMessage());
            }
        }, ['get'], true);
    }
}
