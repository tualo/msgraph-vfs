<?php

namespace Tualo\Office\MSGraphVFS\StreamWrapper;

use GuzzleHttp\Psr7\Utils;
use Microsoft\Graph\Generated\Models\DriveItem;
use Microsoft\Graph\Generated\Models\Folder;
use Microsoft\Graph\GraphServiceClient;
use Psr\Http\Message\StreamInterface;
use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\MSGraph\API;

class SharePointStreamWrapper
{
    private static array $configuration = [];

    private static ?GraphServiceClient $graphClient = null;

    private string $path = '';

    private string $relativePath = '';

    private string $driveItemId = 'root';

    private string $mode = 'r';

    private bool $dirty = false;

    private string $buffer = '';

    private int $position = 0;

    /** @var array<int, string> */
    private array $entries = [];

    private int $entryPosition = 0;

    public static function configure(array $configuration): void
    {
        self::$configuration = $configuration;
    }

    private static function getConfig(string $key, mixed $default = null): mixed
    {
        if (array_key_exists($key, self::$configuration)) {
            return self::$configuration[$key];
        }

        return App::configuration('msgraph-vfs', $key, $default);
    }

    private static function getGraphClient(): GraphServiceClient
    {
        if (self::$graphClient === null) {
            self::$graphClient = API::GraphClient();
        }

        return self::$graphClient;
    }

    private static function normalizeRelativePath(string $path): string
    {
        $decoded = rawurldecode($path);
        $decoded = preg_replace('#/+#', '/', $decoded) ?? $decoded;
        $decoded = trim($decoded, '/');

        return $decoded;
    }

    private static function parseRelativePath(string $path): string
    {
        $parts = parse_url($path);
        $segments = [];

        if (is_array($parts)) {
            if (isset($parts['host']) && $parts['host'] !== '') {
                $segments[] = $parts['host'];
            }

            if (isset($parts['path']) && $parts['path'] !== '') {
                $segments[] = ltrim($parts['path'], '/');
            }
        } else {
            $segments[] = $path;
        }

        return self::normalizeRelativePath(implode('/', array_filter($segments, static function ($segment) {
            return $segment !== '';
        })));
    }

    private static function encodeDrivePath(string $relativePath): string
    {
        $relativePath = self::normalizeRelativePath($relativePath);

        if ($relativePath === '') {
            return 'root';
        }

        $segments = array_map(static function (string $segment): string {
            return rawurlencode($segment);
        }, explode('/', $relativePath));

        return 'root:/' . implode('/', $segments) . ':';
    }

    private static function resolveDriveId(): string
    {
        $driveId = trim((string) self::getConfig('driveId', ''));
        if ($driveId !== '') {
            return $driveId;
        }

        $siteId = trim((string) self::getConfig('siteId', ''));
        if ($siteId !== '') {
            $drive = self::getGraphClient()->sites()->bySiteId($siteId)->drive()->get()->wait();
            if ($drive !== null && method_exists($drive, 'getId') && $drive->getId() !== null) {
                return $drive->getId();
            }
        }

        $drive = self::getGraphClient()->me()->drive()->get()->wait();
        if ($drive !== null && method_exists($drive, 'getId') && $drive->getId() !== null) {
            return $drive->getId();
        }

        throw new \RuntimeException('No MS Graph drive configured. Set msgraph-vfs.driveId or msgraph-vfs.siteId.');
    }

    private static function resolveSiteIdFromDriveId(string $driveId): ?string
    {
        $drive = self::getGraphClient()->drives()->byDriveId($driveId)->get()->wait();

        if ($drive !== null && method_exists($drive, 'getSharePointIds') && $drive->getSharePointIds() !== null) {
            return $drive->getSharePointIds()->getSiteId();
        }

        return null;
    }

    public static function getDriveId(): string
    {
        return self::resolveDriveId();
    }

    public static function getSiteId(): ?string
    {
        $siteId = trim((string) self::getConfig('siteId', ''));
        if ($siteId !== '') {
            return $siteId;
        }

        $driveId = trim((string) self::getConfig('driveId', ''));
        if ($driveId !== '') {
            return self::resolveSiteIdFromDriveId($driveId);
        }

        try {
            return self::resolveSiteIdFromDriveId(self::resolveDriveId());
        } catch (\Throwable $throwable) {
            return null;
        }
    }

    private static function getDriveItemId(string $relativePath): string
    {
        return self::encodeDrivePath($relativePath);
    }

    private static function getItem(string $driveItemId)
    {
        return self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($driveItemId)->get()->wait();
    }

    private static function getChildren(string $driveItemId): array
    {
        $response = self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($driveItemId)->children()->get()->wait();

        if ($response === null || !method_exists($response, 'getValue') || $response->getValue() === null) {
            return [];
        }

        return $response->getValue();
    }

    private static function getContent(string $driveItemId): string
    {
        $response = self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($driveItemId)->content()->get()->wait();

        if ($response instanceof StreamInterface) {
            return (string) $response;
        }

        if (is_object($response) && method_exists($response, 'getBody')) {
            $body = $response->getBody();
            if ($body instanceof StreamInterface) {
                return (string) $body;
            }
        }

        return (string) $response;
    }

    private static function uploadContent(string $driveItemId, string $content)
    {
        return self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($driveItemId)->content()->put(Utils::streamFor($content))->wait();
    }

    private static function deleteItem(string $driveItemId): void
    {
        self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($driveItemId)->delete()->wait();
    }

    private static function createFolder(string $parentDriveItemId, string $name)
    {
        $folderItem = new DriveItem();
        $folderItem->setName($name);
        $folderItem->setFolder(new Folder());

        return self::getGraphClient()->drives()->byDriveId(self::resolveDriveId())->items()->byDriveItemId($parentDriveItemId)->children()->post($folderItem)->wait();
    }

    private static function toStatArray(int $mode, int $size, int $mtime): array
    {
        return [
            0 => 0,
            1 => 0,
            2 => $mode,
            3 => 1,
            4 => 0,
            5 => 0,
            6 => 0,
            7 => $size,
            8 => $mtime,
            9 => $mtime,
            10 => $mtime,
            11 => -1,
            12 => -1,
            'dev' => 0,
            'ino' => 0,
            'mode' => $mode,
            'nlink' => 1,
            'uid' => 0,
            'gid' => 0,
            'rdev' => 0,
            'size' => $size,
            'atime' => $mtime,
            'mtime' => $mtime,
            'ctime' => $mtime,
            'blksize' => 4096,
            'blocks' => (int) ceil(max($size, 1) / 512),
        ];
    }

    private static function itemToStat($item): array
    {
        $isDirectory = method_exists($item, 'getFolder') && $item->getFolder() !== null;
        $size = 0;

        if (!$isDirectory && method_exists($item, 'getSize') && $item->getSize() !== null) {
            $size = (int) $item->getSize();
        }

        $mtime = time();
        if (method_exists($item, 'getLastModifiedDateTime') && $item->getLastModifiedDateTime() !== null) {
            $mtime = $item->getLastModifiedDateTime()->getTimestamp();
        }

        return self::toStatArray($isDirectory ? 040755 : 0100644, $size, $mtime);
    }

    private function setPath(string $path): void
    {
        $this->path = $path;
        $this->relativePath = self::parseRelativePath($path);
        $this->driveItemId = self::getDriveItemId($this->relativePath);
    }

    public function stream_open($path, $mode, $options, &$opened_path)
    {
        $this->setPath($path);
        $this->mode = (string) $mode;
        $opened_path = $path;

        $writeMode = strpbrk($this->mode, 'waxc+') !== false;
        $appendMode = str_contains($this->mode, 'a');

        try {
            if ($writeMode) {
                $exists = true;
                try {
                    self::getItem($this->driveItemId);
                } catch (\Throwable $throwable) {
                    $exists = false;
                }

                if (!$exists && str_contains($this->mode, 'x')) {
                    return false;
                }

                if ($exists && !$appendMode && !str_contains($this->mode, 'w') && !str_contains($this->mode, 'c') && !str_contains($this->mode, '+')) {
                    return false;
                }

                if ($exists && $appendMode) {
                    $this->buffer = self::getContent($this->driveItemId);
                    $this->position = strlen($this->buffer);
                } elseif ($exists && !str_contains($this->mode, 'w')) {
                    $this->buffer = self::getContent($this->driveItemId);
                    $this->position = 0;
                } else {
                    $this->buffer = '';
                    $this->position = 0;
                }

                $this->dirty = true;
                return true;
            }

            $this->buffer = self::getContent($this->driveItemId);
            $this->position = 0;

            return true;
        } catch (\Throwable $throwable) {
            if (($options & STREAM_REPORT_ERRORS) === STREAM_REPORT_ERRORS) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function stream_read($count)
    {
        $chunk = substr($this->buffer, $this->position, $count);
        $this->position += strlen($chunk);

        return $chunk;
    }

    public function stream_write($data)
    {
        $before = substr($this->buffer, 0, $this->position);
        $after = substr($this->buffer, $this->position + strlen($data));
        $this->buffer = $before . $data . $after;
        $this->position += strlen($data);
        $this->dirty = true;

        return strlen($data);
    }

    public function stream_tell()
    {
        return $this->position;
    }

    public function stream_eof()
    {
        return $this->position >= strlen($this->buffer);
    }

    public function stream_seek($offset, $whence)
    {
        $length = strlen($this->buffer);

        if ($whence === SEEK_SET) {
            $newPosition = $offset;
        } elseif ($whence === SEEK_CUR) {
            $newPosition = $this->position + $offset;
        } elseif ($whence === SEEK_END) {
            $newPosition = $length + $offset;
        } else {
            return false;
        }

        if ($newPosition < 0) {
            return false;
        }

        $this->position = $newPosition;

        return true;
    }

    public function stream_flush()
    {
        if (!$this->dirty || strpbrk($this->mode, 'waxc+') === false) {
            return true;
        }

        try {
            self::uploadContent($this->driveItemId, $this->buffer);
            $this->dirty = false;

            return true;
        } catch (\Throwable $throwable) {
            if (function_exists('trigger_error')) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function stream_close(): void
    {
        $this->stream_flush();
    }

    public function url_stat($path, $flags)
    {
        try {
            $this->setPath($path);
            $item = self::getItem($this->driveItemId);

            return self::itemToStat($item);
        } catch (\Throwable $throwable) {
            if (($flags & STREAM_URL_STAT_QUIET) !== STREAM_URL_STAT_QUIET) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function stream_stat()
    {
        return $this->url_stat($this->path, 0);
    }

    public function dir_opendir($path, $options)
    {
        try {
            $this->setPath($path);
            $children = self::getChildren($this->driveItemId);
            $this->entries = array_map(static function ($item): string {
                return method_exists($item, 'getName') && $item->getName() !== null ? (string) $item->getName() : '';
            }, $children);
            $this->entries = array_values(array_filter($this->entries, static function (string $entry): bool {
                return $entry !== '';
            }));
            $this->entryPosition = 0;

            return true;
        } catch (\Throwable $throwable) {
            if (($options & STREAM_REPORT_ERRORS) === STREAM_REPORT_ERRORS) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function dir_readdir()
    {
        if (!isset($this->entries[$this->entryPosition])) {
            return false;
        }

        return $this->entries[$this->entryPosition++];
    }

    public function dir_rewinddir()
    {
        $this->entryPosition = 0;

        return true;
    }

    public function dir_closedir()
    {
        $this->entries = [];
        $this->entryPosition = 0;

        return true;
    }

    public function mkdir($path, $mode, $options)
    {
        try {
            $relativePath = self::parseRelativePath($path);
            $relativePath = trim($relativePath, '/');

            if ($relativePath === '') {
                return false;
            }

            $segments = explode('/', $relativePath);
            $folderName = array_pop($segments);
            $parentPath = implode('/', $segments);
            $parentDriveItemId = self::getDriveItemId($parentPath);

            self::createFolder($parentDriveItemId, $folderName);

            return true;
        } catch (\Throwable $throwable) {
            if (($options & STREAM_REPORT_ERRORS) === STREAM_REPORT_ERRORS) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function rmdir($path, $options)
    {
        try {
            $relativePath = self::parseRelativePath($path);
            self::deleteItem(self::getDriveItemId($relativePath));

            return true;
        } catch (\Throwable $throwable) {
            if (($options & STREAM_REPORT_ERRORS) === STREAM_REPORT_ERRORS) {
                trigger_error($throwable->getMessage(), E_USER_WARNING);
            }

            return false;
        }
    }

    public function unlink($path)
    {
        try {
            self::deleteItem(self::getDriveItemId(self::parseRelativePath($path)));

            return true;
        } catch (\Throwable $throwable) {
            trigger_error($throwable->getMessage(), E_USER_WARNING);

            return false;
        }
    }
}
