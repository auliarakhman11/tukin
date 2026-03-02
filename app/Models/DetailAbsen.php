<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class DetailAbsen extends Model
{
    use HasFactory;

    protected $table = 'detail_absen';
    protected $fillable = ['absensi_id','nip','nm_pegawai','tgl','jam_masuk','jam_keluar','scan_masuk','scan_keluar','denda_masuk','denda_pulang','potongan','jenis','ket','user_id'];
}
