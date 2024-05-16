<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel\Providers;

use Mini\Support\ServiceProvider;
use MiniExcel\Excel;

class MiniExcelServiceProvider extends ServiceProvider
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