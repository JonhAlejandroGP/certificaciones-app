<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\FormularioController;
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
    return view('Formulario');
});
Route::get('/buscar', [FormularioController::class, 'buscar']);
Route::post('/subir', [FormularioController::class, 'procesarArchivo']);
