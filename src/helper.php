<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel;

use Mini\Support\Collection;

if (!function_exists('excel')) {
    /**
     * Return app instance of FastExcel.
     * @param mixed $data
     * @return Excel
     */
    function excel(\Generator|Collection|array $data = null): Excel
    {
        if ($data instanceof Collection) {
            return app('mini.excel')->data($data);
        }

        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        return $data === null ? app('mini.excel') : app('mini.excel')->data($data);
    }
}
