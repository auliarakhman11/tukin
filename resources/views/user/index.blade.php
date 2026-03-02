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

            <div class="col-8">
                <div class="card">
                  {{-- @if (session('success'))
                  <div class="alert alert-success">
                      {{ session('success') }}
                  </div>
                  @endif --}}
                    <div class="card-header">
                        <h4 class="float-left">Data User</h4>
                        <button type="button" id="btn-add-user" class="btn btn-sm btn-primary float-right" data-toggle="modal" data-target="#modal-add-user">
                            <i class="fas fa-plus-circle"></i> 
                            Tambah User
                        </button>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table class="table table-sm" id="table_users">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Nama</th>
                                        <th>Username</th>
                                        <th>Seksi</th>
                                        <th>Aksi</th>
                                    </tr>
                                </thead>
                                <tbody>
                                  
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>


        </div>
      </div><!--/. container-fluid -->
    </section>
    <!-- /.content -->
  </div>
  <!-- /.content-wrapper -->

  <!-- Modal -->
  <form id="form-add-user">
    <div class="modal fade" id="modal-add-user" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header bg-primary">
            <h5 class="modal-title" id="exampleModalLabel">Tambah User</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                <span aria-hidden="true">&times;</span>
            </button>
            </div>
            <div class="modal-body">
                <div id="message_error" style="display:none"></div>
                <div class="row">
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Nama</label>
                            <input type="text" class="form-control" name="name" required>
                        </div>
                    </div>
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Username</label>
                            <input type="text" class="form-control" name="username" required>
                        </div>
                    </div>
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Seksi</label>
                            <select name="seksi_id" class="form-control select2bs4" required>
                                <option value="" >-Pilih Seksi-</option>
                                @foreach ($seksi as $o)
                                <option value="{{ $o->id }}" >{{ $o->nm_seksi }}</option> 
                                @endforeach
                            </select>
                        </div>
                    </div>
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Password</label>
                            <input type="password" class="form-control" name="password" required>
                        </div>
                    </div>
                    <div class="col-12">
                        <div class="form-group">
                            <label for="">Ulangi Password</label>
                            <input type="password" class="form-control" name="password_confirmation" required>
                        </div>
                    </div>
                </div>
            </div>
            <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
            <button type="submit" class="btn btn-primary" id="btn-input-user">Tambah</button>
            </div>
        </div>
        </div>
    </div>
    </form>

    <form id="form_edit_user">
        <div class="modal fade" id="modal_edit_user" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header bg-primary">
                <h5 class="modal-title" id="exampleModalLabel">Edit User</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <input type="hidden" name="id" id="id_e">
                        <div class="col-12">
                            <div class="form-group">
                                <label for="">Nama</label>
                                <input type="text" class="form-control" name="name" id="name_e" required>
                            </div>
                        </div>

                        <div class="col-12">
                            <div class="form-group">
                                <label for="">Seksi</label>
                                <select name="seksi_id" class="form-control select2bs4" id="seksi_id_e" required>
                                    @foreach ($seksi as $o)
                                    <option value="{{ $o->id }}" >{{ $o->nm_seksi }}</option> 
                                    @endforeach
                                </select>
                            </div>
                        </div>

                    </div>
                </div>
                <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                <button type="submit" class="btn btn-primary" id="btn_edit_user">Edit</button>
                </div>
            </div>
            </div>
        </div>
        </form>

    {{-- <form id="form-edit-user">
        <div class="modal fade" id="modal-edit-user" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header bg-primary">
                <h5 class="modal-title" id="exampleModalLabel">Edit Jenis user</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <input type="hidden" name="id" id="id_user_e">
                        <div class="col-12">
                            <div class="form-group">
                                <label for="">Nama Jenis user</label>
                                <input type="text" class="form-control" id="nm_user_e" name="nm_user" required>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                <button type="submit" class="btn btn-primary" id="btn-edit-user">Edit</button>
                </div>
            </div>
            </div>
        </div>
        </form> --}}

          

@section('script')
<script>

$(document).ready(function () {
            $.ajaxSetup({
                headers: {
                    'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                }
            });
        });

  $(document).ready(function() {

//user
    $('#table_users').DataTable({
                processing: true,
                serverSide: true, //aktifkan server-side 
                ajax: {
                    url: "{{ route('getDataUser') }}",
                    type: 'GET'
                },
                columns: [
                    {
                        data: 'DT_RowIndex',
                        name: 'DT_RowIndex'
                    },          
                    {
                        data: 'name',
                        name: 'name'
                    },
                    {
                        data: 'username',
                        name: 'username'
                    },
                    {
                        data: 'seksi.nm_seksi',
                        name: 'seksi'
                    },
                    {
                        data: 'action',
                        name: 'action'
                    }
                ],
                order: [
                    [0, 'asc']
                ]
            });

            $(document).on('submit', '#form-add-user', function(event) {
                event.preventDefault();
                    $('#btn-input-user').attr('disabled',true);
                    $('#btn-input-user').html('Loading <div class="ld"><div></div><div></div><div></div></div>');
                    $('#message_error').hide();
                    $.ajax({
                        url:"{{ route('addUser') }}",
                        method: 'POST',
                        data: new FormData(this),
                        contentType: false,
                        processData: false,
                        success: function(data) {

                                $('#form-add-user').trigger("reset"); //form reset
                                $('#modal-add-user').modal('hide'); //modal hide
                                $("#btn-input-user").removeAttr("disabled");
                                $('#btn-input-user').html('Tambah'); //tombol simpan

                                var oTable = $('#table_users').dataTable(); //inialisasi datatable
                                oTable.fnDraw(false); //reset datatable

                                Swal.fire({
                                toast: true,
                                position: 'top-end',
                                showConfirmButton: false,
                                timer: 3000,
                                icon: 'success',
                                title: 'Data berhasil diinput'
                                });
                            
                                
                                                        
                        },
                        error: function (data) { //jika error tampilkan error pada console
                                    console.log('Error:', data);
                                    var dt_error = '<div class="alert alert-danger">';
                                    jQuery.each(data.responseJSON.errors, function(key, message){
                                        
                                    dt_error += '<p>'+message+'</p>';
                                    // $('.alert-danger').append('<p>'+message+'</p>');
                                    });
                                    dt_error += '</div>';
                                    $('#message_error').html(dt_error);
                                    $('#message_error').show();

                                    $('#btn-input-user').html('Tambah');
                                    $("#btn-input-user").removeAttr("disabled");
                                }
                    });

                });

        $(document).on('click', '.edit_user', function() {
            var id = $(this).data('id');            
            $.get('get-user/' + id, function (data) {
                //set value masing-masing id berdasarkan data yg diperoleh dari ajax get request diatas               
                $('#name_e').val(data.name);
                $('#id_e').val(data.id);

                $('#seksi_id_e').val(data.seksi_id);
                $('#seksi_id_e').select2({theme: 'bootstrap4', tags: true,}).trigger('change');
                
            });
        });


        $(document).on('submit', '#form_edit_user', function(event) {
                event.preventDefault();
                    $('#btn_edit_user').attr('disabled',true);
                    $('#btn_edit_user').html('Loading <div class="ld"><div></div><div></div><div></div></div>');
                    $('#message_error').hide();
                    $.ajax({
                        url:"{{ route('editUser') }}",
                        method: 'POST',
                        data: new FormData(this),
                        contentType: false,
                        processData: false,
                        success: function(data) {

                                $('#form_edit_user').trigger("reset"); //form reset
                                $('#modal_edit_user').modal('hide'); //modal hide
                                $("#btn_edit_user").removeAttr("disabled");
                                $('#btn_edit_user').html('Edit'); //tombol simpan

                                var oTable = $('#table_users').dataTable(); //inialisasi datatable
                                oTable.fnDraw(false); //reset datatable

                                Swal.fire({
                                toast: true,
                                position: 'top-end',
                                showConfirmButton: false,
                                timer: 3000,
                                icon: 'success',
                                title: 'Data berhasil diedit'
                                });
                            
                                
                                                        
                        },
                        error: function (data) { //jika error tampilkan error pada console
                                    console.log('Error:', data);

                                    $('#btn_edit_user').html('Edit');
                                    $("#btn_edit_user").removeAttr("disabled");
                                }
                    });

                });

    //end user

    
    
  });

</script>
@endsection
@endsection  
  
