<?php

use App\Http\Controllers\AuthController;
use App\Http\Controllers\TukinController;
use App\Http\Controllers\UserController;

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

// Route::get('/', function () {
//     return view('welcome');
// });

Route::middleware('auth')->group(function(){

    Route::get('/',[TukinController::class,'index'])->name('tukin');
    Route::post('importAbsen',[TukinController::class,'importAbsen'])->name('importAbsen');
    Route::post('deleteAbsensi',[TukinController::class,'deleteAbsensi'])->name('deleteAbsensi');
    Route::get('exportAbsen/{id}',[TukinController::class,'exportAbsen'])->name('exportAbsen');

    Route::get('/logout',[AuthController::class,'logout'])->name('logout');

    Route::middleware('hakakses:1')->group(function(){

        //user
        Route::get('user',[UserController::class,'index'])->name('user');
        Route::get('get-data-user',[UserController::class,'getDataUser'])->name('getDataUser');
        Route::post('user',[UserController::class,'addUser'])->name('addUser');

        Route::get('get-user/{id}',[UserController::class,'getUser'])->name('getUser');

        Route::post('edit-user',[UserController::class,'editUser'])->name('editUser');
        //enduser

    });
    

    //block
    Route::get('forbidden-access',[AuthController::class,'block'])->name('block');
    //endblock

    Route::get('ganti-password',[UserController::class,'gantiPassword'])->name('gantiPassword');

    Route::post('edit-password',[UserController::class,'editPassword'])->name('editPassword');

   

});


Route::middleware('guest')->group(function(){
    Route::get('login',[AuthController::class,'login_page'])->name('loginPage');
    Route::post('login',[AuthController::class,'login'])->name('login');
});



