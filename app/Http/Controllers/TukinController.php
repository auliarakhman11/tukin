<?php

namespace App\Http\Controllers;

use App\Models\Absensi;
use App\Models\DetailAbsen;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class TukinController extends Controller
{
    public function index()
    {
        return view('tukin.index',[
            'title' => 'Tukin',
            'absensi' => Absensi::orderBy('id','DESC')->with('user')->get(),
        ]);
    }

    public function deleteAbsensi(Request $reader)
    {
        Absensi::where('id',$reader->id)->delete();
        DetailAbsen::where('absensi_id',$reader->id)->delete();

        return redirect()->back()->with('success','Data berhasil dihapus');
    }

    public function isWeekend($date)
    {
        if(date('N', strtotime($date)) ==5){
            return 'jumat';
        }elseif(date('N', strtotime($date)) >= 6){
            return 'weekend';
        }else{
            return 'weekday';
        }
    }

    public function importAbsen(Request $request) {
        $dt_absensi = Absensi::create([
            'tgl' => $request->tahun.'-'.$request->bulan.'-01',
            'user_id' => Auth::user()->id,
        ]);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx;
        $spreadsheet = $reader->load($_FILES['file_excel']['tmp_name']);
        
        $sheet = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);


        $data = array();
        $numrow = 1;

        foreach ($sheet as $row) {

                if ($row['A'] == "" &&  $row['B'] == "" && $row['C'] == "" && $row['D'] == "" && $row['E'] == "" && $row['F'] == "" && $row['G'] == "" && $row['H'] == "")
                    continue;

                // $datetime = DateTime::createFromFormat('Y-m-d', $row['A']);
                if ($numrow > 1) {

                    // $jam_masuk = 'Jebol';
                    // $jam_keluar = 'Jebol';
                    // $scan_masuk = 'Jebol';
                    // $scan_keluar = 'Jebol';
                    // $denda_masuk = 'Jebol';
                    // $denda_pulang = 'Jebol';
                    // $potongan = 'Jebol';

                    if ($row['E'] >= '2026-02-19' && $row['E'] <= '2026-03-18') {

                        if($cek_weekend == 'weekend'){
                            continue;
                        }elseif($cek_weekend == 'jumat'){

                            $jam_masuk = '08:00:00';
                            $jam_keluar = '15:30:00';

                            if ($row['F'] == '-') {
                                $denda_masuk = 1.5;
                            }else{
                                $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                if($telat_masuk > 90){
                                    $denda_masuk = 1.5;
                                }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                    $denda_masuk = 1.25;
                                }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                    $denda_masuk = 1;
                                }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                    $denda_masuk = 0.5;
                                }else{
                                    $denda_masuk = 0;
                                }
                            }

                            if ($row['K'] == '-') {
                                $denda_pulang = 1.5;
                            } else {
                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
                                    
                                if($telat_pulang > 90){
                                    $denda_pulang = 1.5;
                                }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                    $denda_pulang = 1.25;
                                }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                    $denda_pulang = 1;
                                }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                    $denda_pulang = 0.5;
                                }else{
                                    $denda_pulang = 0;
                                }
                            }


                        }else{
                            $jam_masuk = '08:00:00';
                            $jam_keluar = '15:00:00';

                            if ($row['F'] == '-') {
                                $denda_masuk = 1.5;
                            }else{
                                $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                if($telat_masuk > 90){
                                    $denda_masuk = 1.5;
                                }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                    $denda_masuk = 1.25;
                                }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                    $denda_masuk = 1;
                                }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                    $denda_masuk = 0.5;
                                }else{
                                    $denda_masuk = 0;
                                }
                            }

                            if ($row['K'] == '-') {
                                $denda_pulang = 1.5;
                            } else {
                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
                                    
                                if($telat_pulang > 90){
                                    $denda_pulang = 1.5;
                                }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                    $denda_pulang = 1.25;
                                }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                    $denda_pulang = 1;
                                }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                    $denda_pulang = 0.5;
                                }else{
                                    $denda_pulang = 0;
                                }
                            }
                        }

                    } else {
                    
                        $cek_weekend = $this->isWeekend($row['E']);
                        if($cek_weekend == 'weekend'){
                            continue;
                        }elseif($cek_weekend == 'jumat'){

                            if ($row['F'] != '-') {
                                if ($row['F'] > '07:30:00') {

                                    if ($row['F'] > '08:00:00') {

                                        $jam_masuk = '07:30:00';
                                        $jam_keluar = '16:30:00';

                                        $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                        if($telat_masuk > 90){
                                            $denda_masuk = 1.5;
                                        }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                            $denda_masuk = 1.25;
                                        }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                            $denda_masuk = 1;
                                        }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                            $denda_masuk = 0.5;
                                        }else{
                                            $denda_masuk = 0;
                                        }

                                        if ($row['K'] == '-') {
                                            $denda_pulang = 1.5;
                                        } else {
                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }
                                        

                                    } else {
                                        
                                        if ($row['K'] == '-') {
                                            $jam_masuk = '07:30:00';
                                            $jam_keluar = '16:30:00';

                                            $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                            if($telat_masuk > 90){
                                                $denda_masuk = 1.5;
                                            }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                                $denda_masuk = 1.25;
                                            }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                                $denda_masuk = 1;
                                            }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                                $denda_masuk = 0.5;
                                            }else{
                                                $denda_masuk = 0;
                                            }

                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }else{

                                            $jam_masuk = $row['F'];
                                            $selisih_masuk = floor((strtotime($row['F']) - strtotime('07:30:00')) / 60);
                                            $jam_keluar = date('H:i:s', strtotime('16:30:00' . "+$selisih_masuk minutes"));
                                            $denda_masuk = 0;

                                            if ($row['K'] < $jam_keluar) {
                                                $jam_masuk = '07:30:00';
                                                $jam_keluar = '16:30:00';

                                                $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                                if($telat_masuk > 90){
                                                    $denda_masuk = 1.5;
                                                }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                                    $denda_masuk = 1.25;
                                                }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                                    $denda_masuk = 1;
                                                }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                                    $denda_masuk = 0.5;
                                                }else{
                                                    $denda_masuk = 0;
                                                }

                                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
            
                                                if($telat_pulang > 90){
                                                    $denda_pulang = 1.5;
                                                }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                    $denda_pulang = 1.25;
                                                }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                    $denda_pulang = 1;
                                                }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                    $denda_pulang = 0.5;
                                                }else{
                                                    $denda_pulang = 0;
                                                }


                                            }else{
                                                $denda_pulang = 0;
                                            }

                                        }

                                        
                                        

                                    }
                                    
                                } else {
                                
                                    $jam_masuk = $row['F'];
                                    $denda_masuk = 0;

                                    if ($row['K'] == '-') {
                                        
                                        $jam_keluar = '16:30:00';
                                        $denda_pulang = 1.5;
                                    } else {
                                        $selisih_masuk = floor((strtotime('07:30:00') - strtotime($row['F'])) / 60);
                                        $jam_keluar = date('H:i:s', strtotime('16:30:00' . "-$selisih_masuk minutes"));

                                        if ($row['K'] < $jam_keluar) {
                                            $jam_keluar = '16:30:00';

                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
            
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }else{
                                                $denda_pulang = 0;
                                            }
                                    }
                                    

                                }

                            } else {
                                $jam_masuk = '07:30:00';
                                $jam_keluar = '16:30:00';
                                $denda_masuk = 1.5;

                                if ($row['K'] == '-') {
                                    $denda_pulang = 1.5;
                                } else {
                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                    if($telat_pulang > 90){
                                        $denda_pulang = 1.5;
                                    }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                        $denda_pulang = 1.25;
                                    }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                        $denda_pulang = 1;
                                    }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                        $denda_pulang = 0.5;
                                    }else{
                                        $denda_pulang = 0;
                                    }
                                }

                            }
                            
                        }else{
                        
                            if ($row['F'] != '-') {
                                if ($row['F'] > '08:00:00') {

                                    if ($row['F'] > '08:30:00') {

                                        $jam_masuk = '08:00:00';
                                        $jam_keluar = '16:30:00';

                                        $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                        if($telat_masuk > 90){
                                            $denda_masuk = 1.5;
                                        }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                            $denda_masuk = 1.25;
                                        }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                            $denda_masuk = 1;
                                        }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                            $denda_masuk = 0.5;
                                        }else{
                                            $denda_masuk = 0;
                                        }

                                        if ($row['K'] == '-') {
                                            $denda_pulang = 1.5;
                                        } else {
                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }
                                        

                                    } else {
                                        
                                        if ($row['K'] == '-') {
                                            $jam_masuk = '08:00:00';
                                            $jam_keluar = '16:30:00';

                                            $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                            if($telat_masuk > 90){
                                                $denda_masuk = 1.5;
                                            }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                                $denda_masuk = 1.25;
                                            }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                                $denda_masuk = 1;
                                            }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                                $denda_masuk = 0.5;
                                            }else{
                                                $denda_masuk = 0;
                                            }

                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }else{

                                            $jam_masuk = $row['F'];
                                            $selisih_masuk = floor((strtotime($row['F']) - strtotime('08:00:00')) / 60);
                                            $jam_keluar = date('H:i:s', strtotime('16:30:00' . "+$selisih_masuk minutes"));
                                            $denda_masuk = 0;

                                            if ($row['K'] < $jam_keluar) {
                                                $jam_masuk = '08:00:00';
                                                $jam_keluar = '16:30:00';

                                                $telat_masuk = floor((strtotime($row['F']) - strtotime($jam_masuk)) / 60);
                                                if($telat_masuk > 90){
                                                    $denda_masuk = 1.5;
                                                }elseif($telat_masuk > 60 && $telat_masuk <= 90){
                                                    $denda_masuk = 1.25;
                                                }elseif($telat_masuk > 30 && $telat_masuk <= 60){
                                                    $denda_masuk = 1;
                                                }elseif($telat_masuk > 0 && $telat_masuk <= 30){
                                                    $denda_masuk = 0.5;
                                                }else{
                                                    $denda_masuk = 0;
                                                }

                                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
            
                                                if($telat_pulang > 90){
                                                    $denda_pulang = 1.5;
                                                }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                    $denda_pulang = 1.25;
                                                }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                    $denda_pulang = 1;
                                                }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                    $denda_pulang = 0.5;
                                                }else{
                                                    $denda_pulang = 0;
                                                }


                                            }else{
                                                $denda_pulang = 0;
                                            }

                                        }

                                        
                                        

                                    }
                                    
                                } else {
                                
                                    $jam_masuk = $row['F'];
                                    $denda_masuk = 0;

                                    if ($row['K'] == '-') {
                                        $jam_keluar = '16:30:00';
                                        $denda_pulang = 1.5;
                                    } else {
                                        $selisih_masuk = floor((strtotime('08:00:00') - strtotime($row['F'])) / 60);
                                        $jam_keluar = date('H:i:s', strtotime('16:30:00' . "-$selisih_masuk minutes"));

                                        if ($row['K'] < $jam_keluar) {
                                            $jam_keluar = '16:30:00';

                                            $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
            
                                            if($telat_pulang > 90){
                                                $denda_pulang = 1.5;
                                            }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                                $denda_pulang = 1.25;
                                            }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                                $denda_pulang = 1;
                                            }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                                $denda_pulang = 0.5;
                                            }else{
                                                $denda_pulang = 0;
                                            }
                                        }else{
                                                $denda_pulang = 0;
                                            }
                                    }
                                    

                                }

                            } else {
                                $jam_masuk = '08:00:00';
                                $jam_keluar = '16:30:00';
                                $denda_masuk = 1.5;

                                if ($row['K'] == '-') {
                                    $denda_pulang = 1.5;
                                } else {
                                $telat_pulang = floor((strtotime($jam_keluar) - strtotime($row['K'])) / 60);
        
                                    if($telat_pulang > 90){
                                        $denda_pulang = 1.5;
                                    }elseif($telat_pulang > 60 && $telat_pulang <= 90){
                                        $denda_pulang = 1.25;
                                    }elseif($telat_pulang > 30 && $telat_pulang <= 60){
                                        $denda_pulang = 1;
                                    }elseif($telat_pulang > 0 && $telat_pulang <= 30){
                                        $denda_pulang = 0.5;
                                    }else{
                                        $denda_pulang = 0;
                                    }
                                }

                            }

                        }

                    }
                    

                    

                    $scan_masuk = $row['F'] == '-' ? NULL : $row['F'];
                    $scan_keluar = $row['K'] == '-' ? NULL : $row['K'];

                    $potongan = $denda_masuk + $denda_pulang;

                    $daftar_ket = ['Cuti Karena Alasan Penting','Cuti Tahunan','Dinas [Presensi Tidak Mandatori]','Cuti Sakit','Cuti Melahirkan'];

                    if (!in_array($row['P'], $daftar_ket)) {
                        $ket = NULL;
                    } else {
                        $ket = $row['P'];
                    }
                    

                    
                    DetailAbsen::create([
                        'absensi_id' => $dt_absensi->id,
                        'nip' => $row['B'],
                        'nm_pegawai' => $row['C'],
                        'tgl' => $row['E'],
                        'jam_masuk' => $jam_masuk,
                        'jam_keluar' => $jam_keluar,
                        'scan_masuk' => $scan_masuk,
                        'scan_keluar' => $scan_keluar,
                        'denda_masuk' => $denda_masuk,
                        'denda_pulang' => $denda_pulang,
                        'potongan' => $potongan,
                        'jenis' => 1,
                        'ket' => $ket,
                        'user_id' => Auth::id()
                    ]);
                    
                    
                }
                $numrow++; // Tambah 1 setiap kali looping
            }

        return redirect()->back()->with('success','Data berhasil diimport');

    }

    public function exportAbsen($id) {
        $spreadsheet = new Spreadsheet;

        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setTitle('Data Absensi');
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'NIP');
        $spreadsheet->getActiveSheet()->setCellValue('B1', 'Nama Pegawai');
        $spreadsheet->getActiveSheet()->setCellValue('C1', 'Tanggal');
        $spreadsheet->getActiveSheet()->setCellValue('D1', 'Jam Masuk');
        $spreadsheet->getActiveSheet()->setCellValue('E1', 'Jam Keluar');
        $spreadsheet->getActiveSheet()->setCellValue('F1', 'Scan Masuk');
        $spreadsheet->getActiveSheet()->setCellValue('G1', 'Scan Keluar');
        $spreadsheet->getActiveSheet()->setCellValue('H1', 'Potongan Masuk');
        $spreadsheet->getActiveSheet()->setCellValue('I1', 'Potongan Keluar');
        $spreadsheet->getActiveSheet()->setCellValue('J1', 'Total Potongan');
        $spreadsheet->getActiveSheet()->setCellValue('K1', 'Keterangan');

        $style = array(
            'font' => array(
                'size' => 12
            ),
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ),
            ),
            'alignment' => array(
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ),
        );

        $spreadsheet->getActiveSheet()->getStyle('A1:K1')->applyFromArray($style);


        $spreadsheet->getActiveSheet()->getStyle('A1:K1')->getAlignment()->setWrapText(true);

        $absen = DetailAbsen::where('absensi_id',$id)->get();

        $kolom = 2;
        foreach ($absen as $d) {
            $spreadsheet->setActiveSheetIndex(0);
            $spreadsheet->getActiveSheet()->setCellValue('A' . $kolom, "'".$d->nip);
            $spreadsheet->getActiveSheet()->setCellValue('B' . $kolom, $d->nm_pegawai);
            $spreadsheet->getActiveSheet()->setCellValue('C' . $kolom, $d->tgl);
            $spreadsheet->getActiveSheet()->setCellValue('D' . $kolom, $d->jam_masuk ? substr($d->jam_masuk,0,5) : '');
            $spreadsheet->getActiveSheet()->setCellValue('E' . $kolom, $d->jam_keluar ? substr($d->jam_keluar,0,5) : '');
            $spreadsheet->getActiveSheet()->setCellValue('F' . $kolom, $d->scan_masuk ? substr($d->scan_masuk,0,5) : '');
            $spreadsheet->getActiveSheet()->setCellValue('G' . $kolom, $d->scan_keluar ? substr($d->scan_keluar,0,5) : '');
            $spreadsheet->getActiveSheet()->setCellValue('H' . $kolom, $d->denda_masuk.'%');
            $spreadsheet->getActiveSheet()->setCellValue('I' . $kolom, $d->denda_pulang.'%');
            $spreadsheet->getActiveSheet()->setCellValue('J' . $kolom, $d->potongan.'%');
            $spreadsheet->getActiveSheet()->setCellValue('K' . $kolom, $d->ket);

            $kolom++;
        }

        $batas = $kolom - 1;

            $border_collom = array(
                'borders' => array(
                    'allBorders' => array(
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    ),
                )
            );
        
        $spreadsheet->getActiveSheet()->getStyle('A1:K' . $batas)->applyFromArray($border_collom);

        foreach ($spreadsheet->getActiveSheet()->getColumnIterator() as $column) {
            $spreadsheet->getActiveSheet()->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
         }

        $writer = new Xlsx($spreadsheet);

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="Data Absensi.xlsx"');
        header('Cache-Control: max-age=0');

        $writer->save('php://output');



    }

}
