<?php

use App\Http\Controllers\ExportExcelController;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::get('/download/report_stock_balance_xlsx', [ExportExcelController::class, 'report_stock_balance_xlsx']);
Route::get('/download/report_stock_xlsx', [ExportExcelController::class, 'report_stock_xlsx']);
Route::get('/download/report_use_linen_xlsx', [ExportExcelController::class, 'report_use_linen_xlsx']);

Route::get('/download/report_damage_xlsx', [ExportExcelController::class, 'report_damage_xlsx']);

// Route::get('/download/report_damage_daily_xlsx', [ExportExcelController::class, 'report_damage_daily_xlsx']);
Route::get('/download/report_damage_select_type_xlsx', [ExportExcelController::class, 'report_damage_select_type_xlsx']);
// Route::get('/download/report_damage_monthly_xlsx', [ExportExcelController::class, 'report_damage_monthly_xlsx']);

Route::get('/download/report_report_round_factory_xlsx', [ExportExcelController::class, 'report_report_round_factory_xlsx']);