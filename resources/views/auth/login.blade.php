<!DOCTYPE html>
<html lang="en">
<head>
	<title>ATR-BPN</title>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
<!--===============================================================================================-->	
<link rel="icon" type="image/png" href="{{asset('img')}}/Logo_BPN-KemenATR.png"/>
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/vendor/bootstrap/css/bootstrap.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/fonts/font-awesome-4.7.0/css/font-awesome.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/vendor/animate/animate.css">
<!--===============================================================================================-->	
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/vendor/css-hamburgers/hamburgers.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/vendor/select2/select2.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/css/util.css">
	<link rel="stylesheet" type="text/css" href="{{ asset('auth') }}/css/main.css">
<!--===============================================================================================-->
</head>
<body>
	
	<div class="limiter">
		<div class="container-login100">
			<div class="wrap-login100">
				<div class="login100-pic js-tilt" data-tilt>
					{{-- <img src="" alt="IMG" id="logo"> --}}
					<img src="{{ asset('img') }}/Logo_BPN-KemenATR.png" alt="IMG" id="logo">
					{{-- <img src="{{ asset('img') }}/logo arsip1.png" alt="IMG"> --}}
				</div>

				<form class="login100-form validate-form" method="POST" id="form_login" action="{{ route('login') }}">
					@csrf
					<span class="login100-form-title">
						Form Login
					</span>

					@error('username')
						<div class="alert alert-danger" role="alert">
							{{ $message }}
						</div>
					@enderror
					
					<?php if(session('success')): ?>
					<div class="alert alert-success" role="alert">
						{{ session('success') }}
					</div>
					<?php endif; ?>
					

					<div class="wrap-input100 validate-input">
						<input class="input100" type="text" name="username" placeholder="Username" value="{{ old('username') }}" id="username" required>
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-user" aria-hidden="true"></i>
						</span>
					</div>

					<div class="wrap-input100 validate-input" data-validate = "Password is required">
						<input class="input100" type="password" name="password" placeholder="Password">
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-lock" aria-hidden="true"></i>
						</span>
					</div>
					
					<div class="container-login100-form-btn">
						<button class="login100-form-btn" id="btn_login" type="submit">
							Login
						</button>
					</div>

					{{-- <div class="text-center p-t-12">
						<span class="txt1">
							Forgot
						</span>
						<a class="txt2" href="#">
							Username / Password?
						</a>
					</div> --}}

					<div class="text-center p-t-136">
						{{-- <a class="txt2" href="#">
							Create your Account
							<i class="fa fa-long-arrow-right m-l-5" aria-hidden="true"></i>
						</a> --}}
					</div>
				</form>
			</div>
		</div>
	</div>
	
	

	
<!--===============================================================================================-->	
	<script src="{{ asset('auth') }}/vendor/jquery/jquery-3.2.1.min.js"></script>
<!--===============================================================================================-->
	<script src="{{ asset('auth') }}/vendor/bootstrap/js/popper.js"></script>
	<script src="{{ asset('auth') }}/vendor/bootstrap/js/bootstrap.min.js"></script>
<!--===============================================================================================-->
	<script src="{{ asset('auth') }}/vendor/select2/select2.min.js"></script>
<!--===============================================================================================-->
	<script src="{{ asset('auth') }}/vendor/tilt/tilt.jquery.min.js"></script>
	<script >
		$('.js-tilt').tilt({
			scale: 1.1
		})
	</script>
<!--===============================================================================================-->
	<script src="{{ asset('auth') }}/js/main.js"></script>

	<script>
		$(document).ready(function() {

			$(document).on('submit', '#form_login', function(event) {
                    $('#btn_login').attr('disabled',true);
                    $('#btn_login').html('Loading..');
                    

                });

		});
	</script>

</body>
</html>