@extends('template.master')
@section('content')
      <!-- Content Wrapper. Contains page content -->
  <div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <div class="content-header">
      <div class="container-fluid">
        <div class="row mb-2">
          <div class="col-sm-6">
          </div><!-- /.col -->
          <div class="col-sm-6">
            {{-- <ol class="breadcrumb float-sm-right">
              <li class="breadcrumb-item"><a href="#">Home</a></li>
              <li class="breadcrumb-item active">Dashboard v2</li>
            </ol> --}}
          </div><!-- /.col -->
        </div><!-- /.row -->
      </div><!-- /.container-fluid -->
    </div>
    <!-- /.content-header -->

    <!-- Main content -->
    <section class="content">
      <div class="container-fluid">
        <div class="row justify-content-center">

            <div class="col-12">
                <div class="card">
                  {{-- @if (session('success'))
                  <div class="alert alert-success">
                      {{ session('success') }}
                  </div>
                  @endif --}}
                    <div class="card-header">
                        <h4 class="float-left">Import Data Absen</h4>
                    </div>
                    <div class="card-body">
                        <form action="{{ route('importAbsen') }}" method="post" enctype="multipart/form-data">
                            @csrf
                            <div class="row">
                                <div class="col-4">
                                    <div class="form-group">
                                        <label for="">Bulan</label>
                                        <select name="bulan" class="form-control select2bs4" required>
                                            <option value="">Pilih Bulan</option>
                                            <option value="01">Januari</option>
                                            <option value="02">Februari</option>
                                            <option value="03">Maret</option>
                                            <option value="04">April</option>
                                            <option value="05">Mei</option>
                                            <option value="06">Juni</option>
                                            <option value="07">Juli</option>
                                            <option value="08">Agustus</option>
                                            <option value="09">September</option>
                                            <option value="10">Oktober</option>
                                            <option value="11">November</option>
                                            <option value="12">Desember</option>
                                        </select>
                                    </div>
                                </div>

                                <div class="col-3">
                                    <div class="form-group">
                                        <label for="">Tahun</label>
                                        <input type="number" class="form-control" name="tahun" required>
                                    </div>
                                </div>

                                <div class="col-3">
                                    <div class="form-group">
                                        <label for="">File Excel</label>
                                        <input type="file" class="form-control" name="file_excel" accept=".xls, .xlsx, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required>
                                    </div>
                                </div>

                                <div class="col-2 mt-2">
                                <button type="submit" class="btn btn-sm btn-primary mt-4">Import</button>

                            </div>

                            </div>
                        </form>
                    </div>
                </div>
            </div>


        </div>
      </div><!--/. container-fluid -->
    </section>
    <!-- /.content -->

    <!-- Main content -->
    <section class="content">
      <div class="container-fluid">
        <div class="row justify-content-center">

            <div class="col-12">
                <div class="card">
                  {{-- @if (session('success'))
                  <div class="alert alert-success">
                      {{ session('success') }}
                  </div>
                  @endif --}}
                    <div class="card-header">
                        <h4 class="float-left">List Data Import</h4>
                    </div>
                    <div class="card-body">
                        <table class="table" id="table">
                            <thead>
                                <tr>
                                    <th>No</th>
                                    <th>Periode</th>
                                    <th>Admin</th>
                                    <th>Aksi</th>
                                </tr>
                            </thead>
                            <tbody>
                                @php
                                    $i = 1;
                                @endphp
                                @foreach ($absensi as $d)
                                    <tr>
                                        <td>{{ $i++ }}</td>
                                        <td>{{ date("F Y", strtotime($d->tgl)) }}</td>
                                        <td>{{ $d->user->name }}</td>
                                        <td>
                                            <form class="d-inline-block" action="{{ route('deleteAbsensi') }}" method="post">
                                                @csrf
                                                <input type="hidden" name="id" value="{{ $d->id }}">
                                                <button type="submit" onclick="return confirm('Apakah anda yakin ingin menghapus data?')" class="btn btn-xs btn-primary">
                                                <i class="fas fa-trash"></i> 
                                                </button>
                                            </form>
                                            <a href="{{ route('exportAbsen',$d->id) }}" class="btn btn-xs btn-primary"><i class="fas fa-file-excel"></i> Export</a>
                                        </td>
                                    </tr>
                                @endforeach
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>


        </div>
      </div><!--/. container-fluid -->
    </section>
    <!-- /.content -->

  </div>
  <!-- /.content-wrapper -->

  <form id="form-add-jenis">
    <div class="modal fade" id="modal-add-jenis" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header bg-primary">
            <h5 class="modal-title" id="exampleModalLabel">add Jenis Surat</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                <span aria-hidden="true">&times;</span>
            </button>
            </div>
            <div class="modal-body">
                <div class="row">
                    
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Jenis</label>
                            <input type="text" class="form-control" name="nm_jenis" required>
                        </div>
                    </div>

                    <div class="col-12">
                      <div class="form-group">
                          <label for="">Kode</label>
                          <input type="text" class="form-control" name="kode" >
                      </div>
                    </div>

                </div>
            </div>
            <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
            <button type="submit" class="btn btn-primary" id="btn-input-jenis">Tambah</button>
            </div>
        </div>
        </div>
    </div>
    </form>

  <form id="form-edit-jenis">
    <div class="modal fade" id="modal-edit-jenis" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header bg-primary">
            <h5 class="modal-title" id="exampleModalLabel">Edit Jenis Surat</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                <span aria-hidden="true">&times;</span>
            </button>
            </div>
            <div class="modal-body">
                <div class="row">
                    <input type="hidden" name="id" id="id_jenis">
                    
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Jenis</label>
                            <input type="text" class="form-control" name="nm_jenis" id="nm_jenis_e" required>
                        </div>
                    </div>

                    <div class="col-12">
                      <div class="form-group">
                          <label for="">Kode</label>
                          <input type="text" class="form-control" name="kode" id="kode_e" >
                      </div>
                    </div>

                </div>
            </div>
            <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
            <button type="submit" class="btn btn-primary" id="btn-edit-jenis">Edit</button>
            </div>
        </div>
        </div>
    </div>
    </form>

          

@section('script')
<script>

$(document).ready(function () {
            $.ajaxSetup({
                headers: {
                    'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                }
            });

            <?php if(session('success')): ?>
            Swal.fire({
                toast: true,
                position: 'top-end',
                showConfirmButton: false,
                timer: 3000,
                icon: 'success',
                title: '<?= session('success'); ?>'
            });        
            <?php endif; ?>

            <?php if(session('error')): ?>
            Swal.fire({
                toast: true,
                position: 'top-end',
                showConfirmButton: false,
                timer: 3000,
                icon: 'error',
                title: '<?= session('error'); ?>'
            });        
            <?php endif; ?>


        });


</script>
@endsection
@endsection  
  
