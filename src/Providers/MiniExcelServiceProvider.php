<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel\Providers;

use Mini\Service\AbstractServiceProvider;
use MiniExcel\Excel;

class MiniExcelServiceProvider extends AbstractServiceProvider
{
    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot(): void
    {
        //
    }

    /**
     * Register any application services.
     * @return void
     */
    public function register(): void
    {
        $this->app->bind('mini.excel', function () {
            return new Excel();
        });
    }
}
