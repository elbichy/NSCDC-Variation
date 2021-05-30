@extends('layouts.app', ['title' => 'All variation Records'])

@section('content')

    <div class="my-content-wrapper">
        <div class="content-container">
            <div class="sectionWrap">
                {{-- HEADING --}}
                <h6 class="center sectionHeading">All PERSONNEL - (VARIATION)</h6>

                {{-- TABLE --}}
                <div class="sectionTableWrap z-depth-1">
                    <div class="topMenuWrap" style="display: flex; justify-content:space-between; margin-bottom: 20px;">
                        <div class="left">
                            <a href="{{ route('import_data') }}" class="greenBtn btn btn-small green darken-2 left">
                                <i class="fas fa-file-excel right"></i></i> IMPORT FROM EXCELL
                            </a>
                        </div>

                        <button id="ippisBtn" class="ippisBtn btn green btn-small">
                            <i class="fas fa-file-excel right"></i> GENERATE IPPIS TRANS.
                        </button>

                        <button id="financeBtn" class="financeBtn btn green btn-small">
                            <i class="fas fa-file-excel right"></i> GENERATE FIN. VARIATION
                        </button>
                        
                        <button id="adminBtn" class="adminBtn btn btn-small">
                            <i class="fas fa-file-word right"></i> GENERATE ADMIN VARIATION
                        </button>
                    </div>
                    <table class="table centered table-bordered striped highlight" id="users-table">
                        <thead>
                            <tr>
                                <th><input type='checkbox' class='browser-default selectAll'></th>
                                <th>#</th>
                                <th>Service No.</th>
                                <th>IPPIS No.</th>
                                <th>Fullname</th>
                                <th>Rank</th>
                                {{-- <th>Date of 1st Appt.</th> --}}
                                <th>Remark</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tfoot>
                            <tr>
                                <th></th>
                                {{-- <th>#</th> --}}
                                <th>Service No.</th>
                                <th>IPPIS No.</th>
                                <th>Fullname</th>
                                <th>Rank</th>
                                {{-- <th>Date of 1st Appt.</th> --}}
                                <th>Remark</th>
                                <th></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>
        <div class="footer z-depth-1">
            <p>&copy; Nigeria Security & Civil Defence Corps</p>
        </div>
    </div>
@endsection

@push('scripts')
    <script src="{{ asset('js/datatable/jquery.dataTables.min.js') }}"></script>
    <script src="{{ asset('js/datatable/dataTables.buttons.min.js') }}"></script>
    <script src="{{ asset('js/datatable/buttons.flash.min.js') }}"></script>
    <script src="{{ asset('js/datatable/jszip.min.js') }}"></script>
    <script src="{{ asset('js/datatable/pdfmake.min.js') }}"></script>
    <script src="{{ asset('js/datatable/vfs_fonts.js') }}"></script>
    <script src="{{ asset('js/datatable/buttons.html5.min.js') }}"></script>
    <script src="{{ asset('js/datatable/buttons.print.min.js') }}"></script>

    <script>

        $(function() {
            // GENERATE BULK ADMIN VARIATION
            $(document).on('click', '#adminBtn', function() {
                let id = [];
                if (confirm('Are you sure you want to generate variation advice for the selected personnel(s)?')) {
                    $('.personnelCheckbox:checked').each(function() {
                        id.push($(this).val())
                    });
                    if (id.length > 0) {
                        $('.adminBtn').prop('disabled', true).html('PROCESSING...');
                        axios.post(`{!! route('generate_bulk_admin_variation') !!}`, { candidates: id }, {responseType: 'blob'})
                            .then(function(response) {
                                if(response.status == 200){
                                    if(response.data.size == 0){
                                        alert('The selected personnel have no single progression record')
                                        $('.adminBtn').prop('disabled', false).html(`<i class="material-icons right">format_list_bulleted</i> GENERATE ADMIN VARIATION `);
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }else{
                                        $('.adminBtn').prop('disabled', false).html(`<i class="material-icons right">format_list_bulleted</i> GENERATE ADMIN VARIATION`);
                                        const url = window.URL.createObjectURL(new Blob([response.data]));
                                        const link = document.createElement('a');
                                        link.href = url;
                                        link.setAttribute('download', 'bulk_variation_advice.docx');
                                        document.body.appendChild(link);
                                        link.click();
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }
                                }
                            });
                    } else {
                        alert('You must select at least one personnel!');
                    }
                }
            });

            // GENERATE BULK IPPIS TRANSLATION
            $(document).on('click', '#ippisBtn', function() {
                let id = [];
                if (confirm('Are you sure you want to generate IPPIS Trans. for the selected personnel(s)?')) {
                    $('.personnelCheckbox:checked').each(function() {
                        id.push($(this).val())
                    });
                    if (id.length > 0) {
                        $('.ippisBtn').prop('disabled', true).html('PROCESSING...');
                        axios.post(`{!! route('generate_bulk_ippis_translation') !!}`, { candidates: id }, {responseType: 'blob'})
                            .then(function(response) {
                                if(response.status == 200){
                                    if(response.data.size == 0){
                                        alert('The selected personnel have no single progression record')
                                        $('.ippisBtn').prop('disabled', false).html(`<i class="fas fa-file-excel right"></i> IMPORT IPPIS TRANS.`);
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }else{
                                        $('.ippisBtn').prop('disabled', false).html(`<i class="fas fa-file-excel right"></i> IMPORT IPPIS TRANS.`);
                                        const url = window.URL.createObjectURL(new Blob([response.data]));
                                        const link = document.createElement('a');
                                        link.href = url;
                                        link.setAttribute('download', 'ippis_translation.xlsx');
                                        document.body.appendChild(link);
                                        link.click();
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }
                                }
                            });
                    } else {
                        alert('You must select at least one personnel!');
                    }
                }
            });

            // GENERATE BULK FINANCE VARIATION
            $(document).on('click', '#financeBtn', function() {
                let id = [];
                if (confirm('Are you sure you want to generate variation advice for the selected personnel(s)?')) {
                    $('.personnelCheckbox:checked').each(function() {
                        id.push($(this).val())
                    });
                    if (id.length > 0) {
                        $('.financeBtn').prop('disabled', true).html('PROCESSING...');
                        axios.post(`{!! route('generate_bulk_finance_variation') !!}`, { candidates: id }, {responseType: 'blob'})
                            .then(function(response) {
                                if(response.status == 200){
                                    if(response.data.size == 0){
                                        alert('The selected personnel have no single progression record')
                                        $('.financeBtn').prop('disabled', false).html(`<i class="fas fa-file-excel right"></i> GENERATE FIN. VARIATION`);
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }else{
                                        $('.financeBtn').prop('disabled', false).html(`<i class="fas fa-file-excel right"></i> GENERATE FIN. VARIATION`);
                                        const url = window.URL.createObjectURL(new Blob([response.data]));
                                        const link = document.createElement('a');
                                        link.href = url;
                                        link.setAttribute('download', 'bulk_variation_advice.xlsx');
                                        document.body.appendChild(link);
                                        link.click();
                                        $('#users-table th input:checked'). prop("checked", false);
                                        $('#users-table').DataTable().ajax.reload();
                                    }
                                }
                            });
                    } else {
                        alert('You must select at least one personnel!');
                    }
                }
            });

            $(document).on('change', '.selectAll', function() {
                if (this.checked) {
                    $('.personnelCheckbox').attr('checked', true);
                } else {
                    $('.personnelCheckbox').attr('checked', false);
                }
            });

            $('#users-table').DataTable({
                dom: 'lBfrtip',
                buttons: [
                    'csv', 'excel', 'pdf'
                ],
                "lengthMenu": [[10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, -1], [10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, "All"]],
                processing: true,
                serverSide: true,
                ajax:  `{!! route('get_all') !!}`,
                columns: [
                    { data: 'checkbox', name: 'checkbox', orderable: false, searchable: false},
                    {
                        "data": "id",
                        "title": "#",
                        render: function (data, type, row, meta) {
                            return meta.row + meta.settings._iDisplayStart + 1;
                        }, "orderable": false, "searchable": false
                    },
                    { data: 'svc_no', name: 'svc_no'},
                    { data: 'ippis_no', name: 'ippis_no'},
                    { data: 'name', name: 'name' },
                    { data: 'present_gl', name: 'present_gl' },
                    // { data: 'dofa', name: 'dofa' },
                    { data: 'remark', name: 'remark'},
                    { data: 'view', name: 'view', "orderable": false, "searchable": false},
                ],
                initComplete: function () {
                    this.api().columns().every(function () {
                        var column = this;
                        var input = document.createElement("input");
                        $(input).attr('placeholder', 'Search');
                        $(input).appendTo($(column.footer()).empty())
                        .on('keyup', function () {
                            column.search($(this).val(), false, false, true).draw();
                        });
                    });
                }
            });
            $('.dataTables_length > label > select').addClass('browser-default');
        });

    </script>

@endpush