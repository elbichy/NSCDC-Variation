<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\VariationController;
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



Route::get('/', [VariationController::class, 'all'])->name('manage_all');
Route::get('/variation/all', [VariationController::class, 'all'])->name('manage_all');
Route::get('/variation/get_all', [VariationController::class, 'get_all_all'])->name('get_all');

Route::get('/variation/conpass', [VariationController::class, 'conpass'])->name('manage_conpass');
Route::get('/variation/get_all/conpass', [VariationController::class, 'get_all_conpass'])->name('get_all_conpass');

Route::get('/variation/conmess', [VariationController::class, 'conmess'])->name('manage_conmess');
Route::get('/variation/get_all/conmess', [VariationController::class, 'get_all_conmess'])->name('get_all_conmess');

Route::get('/variation/conhessp', [VariationController::class, 'conhessp'])->name('manage_conhessp');
Route::get('/variation/get_all/conhessp', [VariationController::class, 'get_all_conhessp'])->name('get_all_conhessp');

Route::get('/variation/conhesshn', [VariationController::class, 'conhesshn'])->name('manage_conhesshn');
Route::get('/variation/get_all/conhesshn', [VariationController::class, 'get_all_conhesshn'])->name('get_all_conhesshn');

Route::group(['prefix' => 'variation'], function () {
    Route::get('/import', [VariationController::class, 'import_data'])->name('import_data');

    Route::post('/personnel/store', [VariationController::class, 'store_imported_personnel'])->name('store_imported_personnel');
    Route::post('/variation/store', [VariationController::class, 'store_imported_variation'])->name('store_imported_variation');
});

// GENERATION OF ADMIN VARIATION
Route::post('/generate/variation/admin/bulk', [VariationController::class, 'generate_bulk_admin_variation'])->name('generate_bulk_admin_variation');
Route::get('/generate/variation/admin/{candidate}', [VariationController::class, 'generate_single_admin_variation'])->name('generate_single_admin_variation');

// GENERATION OF FINANCE VARIATION
Route::post('/generate/variation/finance/bulk', [VariationController::class, 'generate_bulk_finance_variation'])->name('generate_bulk_finance_variation');
Route::get('/generate/variation/finance/{candidate}', [VariationController::class, 'generate_single_finance_variation'])->name('generate_single_finance_variation');

// GENERATION OF IPPIS TRANSLATION
Route::post('/generate/variation/finance/ippis', [VariationController::class, 'generate_bulk_ippis_translation'])->name('generate_bulk_ippis_translation');
