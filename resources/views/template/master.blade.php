<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="csrf-token" content="{{ csrf_token() }}">
  <title>{{ $title }} | ATR - BPN</title>

  <link rel="icon" type="image/png" href="{{asset('img')}}/Logo_BPN-KemenATR.png"/>

  <!-- Google Font: Source Sans Pro -->
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,400i,700&display=fallback">
  <!-- Font Awesome Icons -->
  <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/fontawesome-free/css/all.min.css">
  <!-- overlayScrollbars -->
  <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/overlayScrollbars/css/OverlayScrollbars.min.css">
  <!-- Theme style -->
  <link rel="stylesheet" href="{{ asset('adminlte') }}/dist/css/adminlte.min.css">

  <link rel="stylesheet" type="text/css" href="{{ asset('dropify') }}/dist/css/dropify.min.css">

  <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/datatables-bs4/css/dataTables.bootstrap4.min.css">
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/datatables-responsive/css/responsive.bootstrap4.min.css">
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/datatables-buttons/css/buttons.bootstrap4.min.css">

     <!-- iCheck for checkboxes and radio inputs -->
     <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/icheck-bootstrap/icheck-bootstrap.min.css">

     <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/toastr/toastr.min.css">
     <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/sweetalert2-theme-bootstrap-4/bootstrap-4.min.css">

     <!-- Select2 -->
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/select2/css/select2.min.css">
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/select2-bootstrap4-theme/select2-bootstrap4.min.css">
    <!-- Bootstrap4 Duallistbox -->
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/bootstrap4-duallistbox/bootstrap-duallistbox.min.css">
    <!-- BS Stepper -->
    <link rel="stylesheet" href="{{ asset('adminlte') }}/plugins/bs-stepper/css/bs-stepper.min.css">

    <link rel="stylesheet" href="{{ asset('daterangepicker') }}/daterangepicker.css">
    @yield('chart')

    <style>
      .ld {
        display: inline-block;
        position: relative;
        width: 21px;
        height: 21px;
      }
      .ld div {
        display: inline-block;
        position: absolute;
        left: 3px;
        width: 6px;
        background: #fff;
        animation: c 1.2s cubic-bezier(0, 0.5, 0.5, 1) infinite;
      }
      .ld div:nth-child(1) {
        left: 3px;
        animation-delay: -0.24s;
      }
      .ld div:nth-child(2) {
        left: 12px;
        animation-delay: -0.12s;
      }
      .ld div:nth-child(3) {
        left: 21px;
        animation-delay: 0;
      }
      @keyframes c {
        0% {
          top: 3px;
          height: 24px;
        }
        50%, 100% {
          top: 9px;
          height: 12px;
        }
      }

    </style>
</head>
<body class="hold-transition sidebar-mini layout-fixed layout-navbar-fixed layout-footer-fixed" id="ini_body">
<div class="wrapper">

    @include('template._navbar')
    @include('template._sidebar')
    @yield('content')

  <!-- Preloader -->
  <div class="preloader flex-column justify-content-center align-items-center">
    <img class="animation__wobble" src="{{ asset('img') }}/Logo_BPN-KemenATR.png" alt="AdminLTELogo" height="60" width="60">
  </div>

  <!-- Control Sidebar -->
  <aside class="control-sidebar control-sidebar-dark">
    <!-- Control sidebar content goes here -->
  </aside>
  <!-- /.control-sidebar -->

  <!-- Main Footer -->
  <footer class="main-footer">
    <strong>Copyright &copy; 2022 <a href="">Kantor Pertanahan Kabupaten Banjar</a>.</strong>
    All rights reserved.
  </footer>
</div>
<!-- ./wrapper -->

<!-- REQUIRED SCRIPTS -->
<!-- jQuery -->
<script src="{{ asset('adminlte') }}/plugins/jquery/jquery.min.js"></script>
<!-- Bootstrap -->
<script src="{{ asset('adminlte') }}/plugins/bootstrap/js/bootstrap.bundle.min.js"></script>
<!-- overlayScrollbars -->
<script src="{{ asset('adminlte') }}/plugins/overlayScrollbars/js/jquery.overlayScrollbars.min.js"></script>
<!-- AdminLTE App -->
<script src="{{ asset('adminlte') }}/dist/js/adminlte.js"></script>

<!-- PAGE PLUGINS -->
<!-- jQuery Mapael -->
{{-- <script src="{{ asset('adminlte') }}/plugins/jquery-mousewheel/jquery.mousewheel.js"></script>
<script src="{{ asset('adminlte') }}/plugins/raphael/raphael.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/jquery-mapael/jquery.mapael.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/jquery-mapael/maps/usa_states.min.js"></script> --}}
<!-- ChartJS -->
{{-- <script src="{{ asset('adminlte') }}/plugins/chart.js/Chart.min.js"></script> --}}

<!-- AdminLTE for demo purposes -->
<script src="{{ asset('adminlte') }}/dist/js/demo.js"></script>
<!-- AdminLTE dashboard demo (This is only for demo purposes) -->
<script src="{{ asset('adminlte') }}/dist/js/pages/dashboard2.js"></script>

<script type="text/javascript" src="{{ asset('dropify') }}/dist/js/dropify.min.js"></script>

{{-- datatables --}}
<script src="{{ asset('adminlte') }}/plugins/datatables/jquery.dataTables.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-bs4/js/dataTables.bootstrap4.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-responsive/js/dataTables.responsive.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-responsive/js/responsive.bootstrap4.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-buttons/js/dataTables.buttons.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-buttons/js/buttons.bootstrap4.min.js"></script>

<script src="{{ asset('adminlte') }}/plugins/datatables-responsive/js/dataTables.responsive.min.js"></script>
<script src="{{ asset('adminlte') }}/plugins/datatables-responsive/js/responsive.bootstrap4.min.js"></script>

{{-- swal --}}
<script src="{{ asset('adminlte') }}/plugins/sweetalert2/sweetalert2.min.js"></script>

{{-- validation --}}
<script src="{{ asset('js') }}/jquery.validate.min.js"></script>

{{-- select2 --}}
<script src="{{ asset('adminlte') }}/plugins/select2/js/select2.full.min.js"></script>
<script src="{{ asset('moment') }}/moment.min.js"></script>
<script src="{{ asset('daterangepicker') }}/daterangepicker.js"></script>

<script type="text/javascript">
  $(document).ready(function() {
    $('.dropify').dropify({
      messages: {
        default: '<small>Drag and drop</small>',
        replace: 'Ganti',
        remove: 'Hapus',
        error: 'error'
      }
    });

    //date picker
    $(function() {
      $("input[name='tanggal']").daterangepicker({
        opens: 'up',
        drops: 'up'
      }, function(start, end, label) {

      });
      $('#tanggal').daterangepicker();
      $('#tanggal').on('apply.daterangepicker', function(ev, tanggal) {

        document.getElementById("tgl1").value = tanggal.startDate.format('YYYY-MM-DD');
        document.getElementById("tgl2").value = tanggal.endDate.format('YYYY-MM-DD');
      });
    });
    //end date picker

    $('#table2').DataTable({ 
    "bSort": true,
    // "scrollX": true,
    "paging" : true,
    "stateSave" : true,
    "scrollCollapse" : true
    });

    $('#table3').DataTable({ 
    "bSort": true,
    // "scrollX": true,
    "paging" : true,
    "stateSave" : true,
    "scrollCollapse" : true
    });

    $('#table').DataTable({ 
    "bSort": true,
    // "scrollX": true,
    "paging" : true,
    "stateSave" : true,
    "scrollCollapse" : true
    });

    $('#table4').DataTable({ 
    "bSort": true,
    // "scrollX": true,
    "paging" : true,
    "stateSave" : true,
    "scrollCollapse" : true
    });

    $('.select').select2();

        $('.select2bs4').select2({
            theme: 'bootstrap4',
            tags: true,
        });

  //select2 auto focus
  $(document).on('select2:open', () => {
        document.querySelector('.select2-search__field').focus();
      });
  });
</script>
@yield('script')
<script>
  $(document).ready(function() {

    const dr = localStorage.getItem('dark_mode_surat');

    function checkDr(dr){
      if(dr){
        if(dr == 1){
          $('#ini_body').addClass('dark-mode');
          $('#dark_mode').attr('checked','checked');
        }else{
          $('#ini_body').removeClass('dark-mode');
          $('#dark_mode').removeAttr('checked');
        }
      
    }else{
      $('#ini_body').removeClass('dark-mode');
      $('#dark_mode').removeAttr('checked');
    }
    }

    checkDr(dr);    

    $(document).on('change', '#dark_mode', function() {
              if ($(this).is(':checked')) {
                localStorage.setItem('dark_mode_surat', 1);
                $('#ini_body').addClass('dark-mode');
              }else{
                localStorage.setItem('dark_mode_surat', 0);
                $('#ini_body').removeClass('dark-mode');
              }
            });
  });
  
</script>
</body>
</html>
