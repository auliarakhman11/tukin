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
            
            <div class="col-6">
              <form method="POST" action="{{ route('editPassword') }}" id="form-edit-password">
                @csrf
                <div class="card">
                  @if (session('success'))
                  <div class="alert alert-success">
                      {{ session('success') }}
                  </div>
                  @endif
                  @if (session('error'))
                  <div class="alert alert-danger">
                      {{ session('error') }}
                  </div>
                  @endif
                    <div class="card-header">
                        <h5 class="float-left">Ganti Password
                        </h5>
                    </div>
                    <div class="card-body">
                        <div class="form-group">
                            <label for="">Password Sekarang</label>
                            <input type="password" class="form-control" name="old_password" required>
                        </div>

                        <div class="form-group">
                            <label for="">Password Baru</label>
                            <input type="password" class="form-control @error('password') is-invalid @enderror" name="password" required>
                            @error('password')
                              <span class="text-danger">{{ $message }}</span>
                            @enderror
                        </div>

                        <div class="form-group">
                            <label for="">Ulangi Password Baru</label>
                            <input type="password" class="form-control" name="password_confirmation" required>
                        </div>

                    <button type="submit" class="btn btn-primary  float-right" id="btn-edit-password">Edit</button>
                    </div>
                </div>
              </form>
            </div>
          
        </div>
      </div><!--/. container-fluid -->
    </section>
    <!-- /.content -->
  </div>
  <!-- /.content-wrapper -->

    

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


    
  });

</script>
@endsection
@endsection  
  
