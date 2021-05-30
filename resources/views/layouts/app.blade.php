<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <title>
        {{ config('app.name', 'NSCDC Admin Directorate Database') }}@isset($title) - {{ $title }}@endisset
    </title>
    <link rel="shortcut icon" href="{{ asset('storage/fav.png') }}">
    <link href="{{ asset('fontawesome/css/all.css') }}" rel="stylesheet"> <!-- font-awesome -->
    {{-- <link rel="stylesheet" type="text/css" href="{{asset('countdowntimer/src/css/jQuery.countdownTimer.css')}}" /> --}}
    <link rel="stylesheet" href="{{asset('css/material-icons.css')}}">
    <script src="{{asset('js/jquery-3.3.1.min.js')}}"></script>
    <script src="{{asset('js/jquery-ui.min.js')}}"></script>
    <script src="{{asset('js/axios.min.js')}}"></script>
    <script src="{{asset('js/wnoty.js')}}"></script>
    {{-- <script type="text/javascript" src="{{asset('countdowntimer/src/js/jQuery.countdownTimer.js')}}"></script> --}}

    <script src="{{asset('materialize-css/js/materialize.min.js')}}"></script>
    {{-- <link rel="stylesheet" charset="utf-8" href="https://fonts.googleapis.com/icon?family=Material+Icons"> --}}
    <link rel="stylesheet" charset="utf-8" href="{{asset('materialize-css/css/materialize.min.css')}}">

    {{-- <script src="https://unpkg.com/sweetalert/dist/sweetalert.min.js"></script> --}}
    <script type="text/javascript" src="{{asset('js/custom.js')}}"></script>

    <style>
        :root {
            --primary-bg-dark: #164f6b; 
            --primary-bg-mid: #0e75a7; 
            --primary-bg-light: #039be5;  
            
            --primary-trans-bg-dark: #164f6b;
            --primary-trans-bg-light: #039be5;
            
            --green-light: #27a747;
            --green-dark: #2a8841;
            
            --secondary-bg-dark: #8d1003; 
            --secondary-bg-light: #c91e0b; 
            
            --switch-dark: #164f6b; 
            --switch-light: #039be5; 

            --button-dark: #164f6b; 
            --button-light: #039be5;
            --button-secondary: #8d1003;
        }
    </style>
    <link rel="stylesheet" href="{{ asset('css/datatable/jquery.dataTables.min.css') }}">
    <link rel="stylesheet" href="{{ asset('css/datatable/buttons.dataTables.min.css') }}">
    <link rel="stylesheet"  href="{{asset('css/lightbox.css')}}"/>
    <link rel="stylesheet" href="{{asset('css/wnoty.css')}}">
    <link rel="stylesheet" href="{{asset('css/app.css')}}">
    <link rel="stylesheet" href="{{asset('css/small.css')}}">
</head>
<body>
    <div class="app" id="app">
        {{-- Navbar goes here --}}
        <div id="space-for-sidenave" class="navbar-fixed">
            <nav>
                <div class="nav-wrapper">
                    {{-- Show View --}}
                    <a href="#" id="show-side-nav" class="hide-on-med-and-down">
                        <i class="small material-icons white-text">menu</i>
                    </a>

                    {{-- BREADCRUMB --}}
                    <div class="left breadcrumbWrap hide-on-med-and-up">
                        <a href="/{{request()->segment(1)}}/{{request()->segment(2)}}/" class="breadcrumb">{{(request()->segment(2) == '') ? 'CONPASS' : ucfirst(request()->segment(2))}}</a>
                    </div>
                    
                    {{-- BREADCRUMB --}}
                    <div class="left breadcrumbWrap hide-on-small-and-down">
                        
                        <a href="/variation/conpass" class="breadcrumb">VARIATION</a>

                        {{-- <a href="/{{request()->segment(1)}}/{{request()->segment(2)}}/{{request()->segment(3)}}" class="breadcrumb">{{(request()->segment(3) == '') ? 'Dashbord' : ucfirst(request()->segment(3))}}</a> --}}

                        @if(request()->segment(2) != '')
                            <a href="/{{request()->segment(1)}}/{{ request()->segment(2) }}" class="breadcrumb">{{ strtoupper(request()->segment(2)) }}</a>
                        @endif
                    </div>
                    
                    {{-- OTHER MENU RIGHT --}}
                    <a href="#" data-target="slide-out" class="sidenav-trigger hide-on-med-and-up right"><i class="material-icons">menu</i></a>
                    <ul class="right hide-on-med-and-down">
                        {{-- <li class="logOutBtn">
                            <a href="{{ route('logout') }}" onclick="event.preventDefault(); document.getElementById('logout-form').submit();">
                                <i class="material-icons right">power_settings_new</i>
                            </a> --}}
                            <form id="logout-form" action="{{ route('logout') }}" method="POST" style="display: none;">
                                @csrf
                            </form>
                        {{-- </li>
                    </ul> --}}
                    @auth
                    <!-- Dropdown Structure -->
                    <ul id="dropdown1" class="dropdown-content">
                        <li><a href="#"><i class="material-icons left">person</i> Profile</a></li>
                        <li class="divider"></li>
                        <li>
                            <a href="{{ route('logout') }}" onclick="event.preventDefault(); document.getElementById('logout-form').submit();">
                            <i class="material-icons left">power_settings_new</i> Logout
                            </a>
                        </li>
                    </ul>
                    <ul class="right hide-on-small-only">
                         <p style="padding-right:12px;"><a class="dropdown-trigger"  data-target="dropdown1" href="#!"  style="display: inline-block;">{{ auth()->user()->fullname }} <i class="material-icons right">arrow_drop_down</i></a></p>
                    </ul>
                    @endauth
                </div>
            </nav>
        </div>

        {{-- SIDE NAV --}}
        <ul id="slide-out" class="sidenav sidenav-fixed" style="min-height: 100%; display: flex; flex-direction: column;">
            <div class="sideNavContainer">
                {{-- THE RED LOGO AREA --}}
                <li>
                    <div class="user-view">
                        {{-- Hide View --}}
                        <a href="#" id="hide-side-nav" class="hide-on-med-and-down">
                            <i class="small material-icons white-text">close</i>
                        </a>

                        {{-- BUSINESS LOGO --}}
                        <a href="#user"><img class="circle" src="{{asset('storage/nscdclargelogo.png')}}"></a>
                    
                        {{-- BUSINESS NAME --}}
                        <a href="#name"><span class="white-text name">
                            Nigeria Security & Civil Defence Corps
                        </span></a>

                        {{-- BUSINESS BRANCH AND ADDRESS --}}
                        <a href="#email"><span class="white-text email">Admin Database</span></a>
                    </div>
                </li>

                {{-- VARIATION --}}
                <li class="no-padding">
                    <ul class="collapsible collapsible-accordion">
                    <li class="{{ request()->segment(1) == 'variation' ? 'active' :  ''}}">
                        <a style="padding:0 32px;" class="collapsible-header">
                            <i class="fal fa-sort-circle-up fa-2x"></i></i>VARIATION<i class="material-icons right">arrow_drop_down</i>
                        </a>
                        <div class="collapsible-body">
                        <ul>
                            <li class="{{(request()->segment(2) == 'all') ? 'active' : ''}}">
                                <a href="{{ route('manage_all') }}">ALL</a>
                            </li>
                            {{-- <li class="{{(request()->segment(2) == 'conpass') ? 'active' : ''}}">
                                <a href="{{ route('manage_conpass') }}">CONPASS</a>
                            </li>
                            <li class="{{(request()->segment(2) == 'conmess') ? 'active' : ''}}">
                                <a href="{{ route('manage_conmess') }}">CONMESS</a>
                            </li>
                            <li class="{{(request()->segment(2) == 'conhessp') ? 'active' : ''}}">
                                <a href="{{ route('manage_conhessp') }}">CONHESSP</a>
                            </li>
                            <li class="{{(request()->segment(2) == 'conhesshn') ? 'active' : ''}}">
                                <a href="{{ route('manage_conhesshn') }}">CONHESSHN</a>
                            </li> --}}
                            <li class="{{(request()->segment(2) == 'import') ? 'active' : ''}}">
                                <a href="{{ route('import_data') }}">Import records</a>
                            </li>
                        </ul>
                        </div>
                    </li>
                    </ul>
                </li>
            </div>
        </ul>

        {{-- CONTENT AREA    --}}
        @if (session()->has('success'))
            <script>
            $(document).ready(function () {
                    $.wnoty({
                    type: 'success',
                    message: '{{session('success')}}',
                    autohide: true,
                    autohideDelay: 5000
                    });
                });
            </script>
        @endif
        @if (session()->has('error'))
            <script>
            $(document).ready(function () {
                    $.wnoty({
                    type: 'error',
                    message: '{{session('error')}}',
                    autohideDelay: 5000
                    });
                });
            </script>
        @endif
        @yield('content')
    </div>
        @include('sweetalert::alert')
    <script>
        let base_url = '{{ asset('/') }}';

        $(document).ready(function(){
            $('#hide-side-nav').click(function(e){
                $(this).fadeOut();
                $('#slide-out').animate({
                    'width': '0px'
                });
                $('.my-content-wrapper').animate({
                    'padding-left': '0px'
                });
                $('#users-table').animate({
                    'width': '100%'
                });
                $('#users-table').animate({
                    'width': '100%'
                });
                $('.breadcrumbWrap').animate({
                   'margin-left': '0px'
                });
                $('#show-side-nav').animate({
                    'width': '60px',
                    'margin-right': '20px'
                });
            });
            $('#show-side-nav').click(function(e){
                $('#hide-side-nav').fadeIn();
                $('#slide-out').animate({
                    'width': '300px'
                });
                $('.my-content-wrapper').animate({
                    'padding-left': '300px'
                });
                $('#users-table').animate({
                    'width': '100%'
                });
                $('.breadcrumbWrap').animate({
                   'margin-left': '310px'
                });
                $('#show-side-nav').animate({
                    'width': '0px',
                    'margin-right': '0px'
                });
            });
        });
    </script>
    <script src="{{ asset('js/lightbox.js') }}"></script>
    @stack('scripts')
</body>
</html>
