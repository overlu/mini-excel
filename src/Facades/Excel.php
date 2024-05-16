<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel\Facades;

use Mini\Facades\Facade;

/**
 * Class Excel.
 * @method static \MiniExcel\Excel data($data)
 * @method static \Mini\Support\Collection import(string $path, callable $callback = null)
 * @method static string export(string $path, callable $callback = null)
 * @method static \Mini\Support\Collection importSheets(string $path, callable $callback = null)
 * @method static \MiniExcel\Excel configureCsv($delimiter = ',', $enclosure = '"', $encoding = 'UTF-8', $bom = false)
 * @method static \MiniExcel\Excel configureReaderUsing(?callable $callback = null)
 * @method static \MiniExcel\Excel configureWriterUsing(?callable $callback = null)
 * @see \MiniExcel\Excel
 */
class Excel extends Facade
{
    /**
     * Get the registered name of the component.
     *
     * @return string
     */
    protected static function getFacadeAccessor(): string
    {
        return 'mini.excel';
    }
}
