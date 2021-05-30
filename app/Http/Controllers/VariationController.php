<?php

namespace App\Http\Controllers;

use App\Models\Variation;
use App\Models\Beneficiary;
use App\Models\Conhesshn;
use App\Models\Conhessp;
use App\Models\Conmess;
use App\Models\Conpass;
use App\Models\OldConpass;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Carbon\Carbon;
use Carbon\CarbonPeriod;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\File;
use Milon\Barcode\DNS2D;
use Rap2hpoutre\FastExcel\FastExcel;
use RealRashid\SweetAlert\Facades\Alert;
use Yajra\DataTables\Facades\DataTables;

class VariationController extends Controller
{

    public function __construct()
    {
        $this->middleware('auth');
    }
    
    // DISPLAY CONPASS LIST
    public function all(Request $request){
        // $variation = Variation::all();
        return view('variation/all');
    }
    // GET variation LIST
    public function get_all_all(){
        $variarions = Variation::orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')");
        // $beneficiaries = Variation::with(['beneficiary' => function($q){
        //     $q->where('salary_structure', 'CONPASS');
        // }])->orderBy('ippis_no');
        return DataTables::eloquent($variarions)
        ->editColumn('dofa', function ($variation) {
                return Carbon::create($variation->dofa)->format('d/m/Y');
        })
        ->addColumn('view', function($variation) {
            return '
                    <a href="'.route('generate_single_admin_variation', $variation->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
                    <a href="'.route('generate_single_finance_variation', $variation->id).'" style="margin-right:5px;" class="green-text text-darken-3" title="Print variation slip"><i class="fas fa-file-excel fa-lg"></i></a>
                ';
        })
        ->addColumn('checkbox', function($variation) {
            return '<input type="checkbox" name="personnelCheckbox[]" class="personnelCheckbox browser-default" value="'.$variation->id.'" />';
        })
        ->rawColumns(['checkbox', 'view'])
        ->make();
    }

    // DISPLAY CONPASS LIST
    public function conpass(Request $request){
        // $variation = Variation::all();
        return view('variation/conpass');
    }
    // GET variation LIST
    public function get_all_conpass(){
        $variarions = Variation::where('salary_structure', 'CONPASS')->orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')");
        // $beneficiaries = Variation::with(['beneficiary' => function($q){
        //     $q->where('salary_structure', 'CONPASS');
        // }])->orderBy('ippis_no');
        return DataTables::eloquent($variarions)
        ->editColumn('dofa', function ($variation) {
                return Carbon::create($variation->dofa)->format('d/m/Y');
        })
        ->addColumn('view', function($variation) {
            return '
                    <a href="'.route('generate_single_admin_variation', $variation->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
                ';
        })
        ->addColumn('checkbox', function($variation) {
            return '<input type="checkbox" name="personnelCheckbox[]" class="personnelCheckbox browser-default" value="'.$variation->id.'" />';
        })
        ->rawColumns(['checkbox', 'view'])
        ->make();
    }
    
    // DISPLAY CONPASS LIST
    public function conmess(Request $request){
        // $variation = Variation::all();
        return view('variation/conmess');
    }
    // GET variation LIST
    public function get_all_conmess(){
        $variarions = Variation::where('salary_structure', 'CONMESS')->orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')");
        return DataTables::eloquent($variarions)
        ->editColumn('dofa', function ($variarion) {
                return Carbon::create($variarion->dofa)->format('d/m/Y');
        })
        ->addColumn('view', function($variarion) {
            return '
                    <a href="'.route('generate_single_admin_variation', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
                ';
        })
        ->addColumn('checkbox', function($variarion) {
            return '<input type="checkbox" name="personnelCheckbox[]" class="personnelCheckbox browser-default" value="'.$variarion->id.'" />';
        })
        ->rawColumns(['checkbox', 'view'])
        ->make();
    }
    
    // DISPLAY CONPASS LIST
    public function conhessp(Request $request){
        // $variation = Variation::all();
        return view('variation/conhessp');
    }
    // GET variation LIST
    public function get_all_conhessp(){
        $variarions = Variation::where('salary_structure', 'CONHESSP')->orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')");
        return DataTables::eloquent($variarions)
        ->editColumn('dofa', function ($variarion) {
                return Carbon::create($variarion->dofa)->format('d/m/Y');
        })
        ->addColumn('view', function($variarion) {
            return '
                    <a href="'.route('generate_single_admin_variation', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
                ';
        })
        ->addColumn('checkbox', function($variarion) {
            return '<input type="checkbox" name="personnelCheckbox[]" class="personnelCheckbox browser-default" value="'.$variarion->id.'" />';
        })
        ->rawColumns(['checkbox', 'view'])
        ->make();
    }
    
    // DISPLAY CONPASS LIST
    public function conhesshn(Request $request){
        // $variation = Variation::all();
        return view('variation/conhesshn');
    }
    // GET variation LIST
    public function get_all_conhesshn(){
        $variarions = Variation::where('salary_structure', 'CONHESSHN')->orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')");
        return DataTables::eloquent($variarions)
        ->editColumn('dofa', function ($variarion) {
                return Carbon::create($variarion->dofa)->format('d/m/Y');
        })
        ->addColumn('view', function($variarion) {
            return '
                    <a href="'.route('generate_single_admin_variation', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
                ';
        })
        ->addColumn('checkbox', function($variarion) {
            return '<input type="checkbox" name="personnelCheckbox[]" class="personnelCheckbox browser-default" value="'.$variarion->id.'" />';
        })
        ->rawColumns(['checkbox', 'view'])
        ->make();
    }

    // IMPORT NEW DATA
    public function import_data(Request $request){
        return view('variation/import');
    }

    // IMPORT STORE IMPORTED PERSONNEL DATA
    public function store_imported_personnel(Request $request)
    {  
        
        $request->validate([
            'import_file' => 'required'
        ]);

        $path = $request->file('import_file')->getRealPath();
        $data = (new FastExcel)->import($path);
        
        if ($data->count()) {
            $candidates = (new FastExcel)->import($path, function ($line) {
                // Beneficiary
                $line['svc_no'] == '' ? $svc_no = null : $svc_no = $line['svc_no'];
                $line['ippis_no'] == '' ? $ippis_no = null : $ippis_no = $line['ippis_no'];
                $line['name'] == '' ? $name = null : $name = $line['name'];
                $line['gender'] == '' ? $gender = null : $gender = $line['gender'];
                $line['dob'] == '' ? $dob = null : $dob = $line['dob'];
                $line['rank'] == '' ? $rank = null : $rank = $line['rank'];
                $line['gl'] == '' ? $gl = null : $gl = $line['gl'];
                $line['dofa'] == '' ? $dofa = null : $dofa = $line['dofa'];
                $line['salary_structure'] == '' ? $salary_structure = null : $salary_structure = $line['salary_structure'];
                // $line['dor'] == '' ? $dor = null : $dor = $line['dor'];
                $line['paypoint'] == '' ? $paypoint = null : $paypoint = $line['paypoint'];

                $beneficiary = Beneficiary::create([
                    'svc_no' => $svc_no,
                    'ippis_no' => $ippis_no,
                    'name' => $name,
                    'gender' => $gender,
                    'dob' => $dob,
                    'rank' => $rank,
                    'gl' => $gl,
                    'dofa' => $dofa,
                    'salary_structure' => $salary_structure,
                    // 'dor' => $dor,
                    'paypoint' => $paypoint,
                ]);
            });
            Alert::success('Personnel varia$variations imported successfully!', 'Success!')->autoclose(222500);
            return back();
        }

    }

    // IMPORT STORE IMPORTED VARIATION DATA
    public function store_imported_variation(Request $request)
    {  
        
        $request->validate([
            'import_file' => 'required'
        ]);

        $path = $request->file('import_file')->getRealPath();
        $data = (new FastExcel)->import($path);
        
        if($data->count()){
            
            $candidates = (new FastExcel)->import($path, function ($line) {
                
                // Variation
                $line['svc_no'] == '' ? $svc_no = null : $svc_no = $line['svc_no'];
                $line['ippis_no'] == '' ? $ippis_no = null : $ippis_no = $line['ippis_no'];
                $line['name'] == '' ? $name = null : $name = $line['name'];
                $line['gender'] == '' ? $gender = null : $gender = $line['gender'];
                $line['dob'] == '' ? $dob = null : $dob = $line['dob'];
                $line['present_rank'] == '' ? $present_rank = null : $present_rank = $line['present_rank'];
                $line['present_gl'] == '' ? $present_gl = null : $present_gl = $line['present_gl'];
                $line['dofa'] == '' ? $dofa = null : $dofa = $line['dofa'];
                $line['salary_structure'] == '' ? $salary_structure = null : $salary_structure = $line['salary_structure'];
                $line['paypoint'] == '' ? $paypoint = null : $paypoint = $line['paypoint'];
                $line['bank'] == '' ? $bank = null : $bank = $line['bank'];
                $line['acc_no'] == '' ? $acc_no = null : $acc_no = $line['acc_no'];
                // $line['dor'] == '' ? $dor = null : $dor = $line['dor'];

                $line['old_rank'] == '' ? $old_rank = null : $old_rank = $line['old_rank'];
                $line['old_gl'] == '' ? $old_gl = null : $old_gl = $line['old_gl'];
                $line['old_step'] == '' ? $old_step = null : $old_step = $line['old_step'];
            
                $line['new_rank'] == '' ? $new_rank = null : $new_rank = $line['new_rank'];
                $line['new_gl'] == '' ? $new_gl = null : $new_gl = $line['new_gl'];
                $line['new_step'] == '' ? $new_step = null : $new_step = $line['new_step'];
                
                $line['effective'] == '' ? $effective = null : $effective = $line['effective'];
                $line['placed'] == '' ? $placed = null : $placed = $line['placed'];
                $line['months_paid'] == '' ? $months_paid = null : $months_paid = $line['months_paid'];
                
                $line['pro_type'] == '' ? $pro_type = null : $pro_type = $line['pro_type'];
                $line['remark'] == '' ? $remark = null : $remark = $line['remark'];
                
                
                // $beneficiary = Beneficiary::find($ippis_no);

                if ($salary_structure == 'CONPASS') {

                    if((int)Carbon::instance($effective)->format('Y') >= 2019){
                        // VARIATION COMPUTATIONS
                        $old_salary_per_annum = Conpass::where('gl', $old_gl)
                        ->where('step', $old_step)
                        ->first()->salary_per_annum;

                        $new_salary_per_annum = Conpass::where('gl', $new_gl)
                                            ->where('step', $new_step)
                                            ->first()->salary_per_annum;
                        $variation_amount = $new_salary_per_annum - $old_salary_per_annum;
                        $effective = Carbon::instance($effective);
                        $placed = Carbon::instance($placed);
                        $months_owed = $effective->diffInMonths($placed);
                        $arrears = $variation_amount*$months_owed;
                        $variation = Variation::create([
                        'svc_no' => $svc_no,
                        'ippis_no' => $ippis_no,
                        'name' => $name,
                        'gender' => $gender,
                        'dob' => $dob,
                        'present_rank' => $present_rank,
                        'present_gl' => $present_gl,
                        'dofa' => $dofa,
                        'salary_structure' => $salary_structure,
                        'paypoint' => $paypoint,
                        'bank' => $bank,
                        'acc_no' => $acc_no,
                        // 'dor' => $dor,
                        'old_rank' => $old_rank,
                        'old_gl' => $old_gl,
                        'old_step' => $old_step,
                        'old_salary_per_annum' => $old_salary_per_annum,
                        'new_rank' => $new_rank,
                        'new_gl' => $new_gl,
                        'new_step' => $new_step,
                        'new_salary_per_annum' => $new_salary_per_annum,
                        'effective' => $effective,
                        'placed' => $placed,
                        'months_owed' => $months_owed,
                        'months_paid' => $months_paid,
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'pro_type' => $remark,
                        'remark' => $remark,
                        ]);
                    }else{
                        // VARIATION COMPUTATIONS
                        $old_salary_per_annum = OldConpass::where('gl', $old_gl)
                        ->where('step', $old_step)
                        ->first()->salary_per_annum;

                        $new_salary_per_annum = OldConpass::where('gl', $new_gl)
                                            ->where('step', $new_step)
                                            ->first()->salary_per_annum;
                        $variation_amount = $new_salary_per_annum - $old_salary_per_annum;
                        $effective = Carbon::instance($effective);
                        $placed = Carbon::instance($placed);
                        $months_owed = $effective->diffInMonths($placed);
                        $arrears = $variation_amount*$months_owed;
                        $variation = Variation::create([
                        'svc_no' => $svc_no,
                        'ippis_no' => $ippis_no,
                        'name' => $name,
                        'gender' => $gender,
                        'dob' => $dob,
                        'present_rank' => $present_rank,
                        'present_gl' => $present_gl,
                        'dofa' => $dofa,
                        'salary_structure' => $salary_structure,
                        'paypoint' => $paypoint,
                        'bank' => $bank,
                        'acc_no' => $acc_no,
                        // 'dor' => $dor,
                        'old_rank' => $old_rank,
                        'old_gl' => $old_gl,
                        'old_step' => $old_step,
                        'old_salary_per_annum' => $old_salary_per_annum,
                        'new_rank' => $new_rank,
                        'new_gl' => $new_gl,
                        'new_step' => $new_step,
                        'new_salary_per_annum' => $new_salary_per_annum,
                        'effective' => $effective,
                        'placed' => $placed,
                        'months_owed' => $months_owed,
                        'months_paid' => $months_paid,
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'pro_type' => $remark,
                        'remark' => $remark,
                        ]);
                    }
                    

                }else if($salary_structure == 'CONMESS'){

                    $old_salary_per_annum = Conmess::where('conpass_gl', $old_gl)
                                                ->where('conpass_step', $old_step)
                                                ->first()->consolidated_salary_per_annum;

                    $new_salary_per_annum = Conmess::where('conpass_gl', $new_gl)
                                                ->where('conpass_step', $new_step)
                                                ->first()->consolidated_salary_per_annum;
                    $variation_amount = $new_salary_per_annum - $old_salary_per_annum;
                    $effective = Carbon::instance($effective);
                    $placed = Carbon::instance($placed);
                    $months_owed = $effective->diffInMonths($placed);
                    $arrears = $variation_amount*$months_owed;
                    $variation = Variation::create([
                        'svc_no' => $svc_no,
                        'ippis_no' => $ippis_no,
                        'name' => $name,
                        'gender' => $gender,
                        'dob' => $dob,
                        'present_rank' => $present_rank,
                        'present_gl' => $present_gl,
                        'dofa' => $dofa,
                        'salary_structure' => $salary_structure,
                        'paypoint' => $paypoint,
                        'bank' => $bank,
                        'acc_no' => $acc_no,
                        // 'dor' => $dor,
                        'old_rank' => $old_rank,
                        'old_gl' => $old_gl,
                        'old_step' => $old_step,
                        'old_salary_per_annum' => $old_salary_per_annum,
                        'new_rank' => $new_rank,
                        'new_gl' => $new_gl,
                        'new_step' => $new_step,
                        'new_salary_per_annum' => $new_salary_per_annum,
                        'effective' => $effective,
                        'placed' => $placed,
                        'months_owed' => $months_owed,
                        'months_paid' => $months_paid,
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'pro_type' => $remark,
                        'remark' => $remark,
                    ]);

                }else if($salary_structure == 'CONHESSP'){

                    $old_salary_per_annum = Conhessp::where('conpass_gl', $old_gl)
                                                ->where('conpass_step', $old_step)
                                                ->first()->consolidated_salary_per_annum;

                    $new_salary_per_annum = Conhessp::where('conpass_gl', $new_gl)
                                                ->where('conpass_step', $new_step)
                                                ->first()->consolidated_salary_per_annum;
                    $variation_amount = $new_salary_per_annum - $old_salary_per_annum;
                    $effective = Carbon::instance($effective);
                    $placed = Carbon::instance($placed);
                    $months_owed = $effective->diffInMonths($placed);
                    $arrears = $variation_amount*$months_owed;
                    $variation = Variation::create([
                        // 'sn' => $sn,
                        'svc_no' => $svc_no,
                        'ippis_no' => $ippis_no,
                        'name' => $name,
                        'gender' => $gender,
                        'dob' => $dob,
                        'present_rank' => $present_rank,
                        'present_gl' => $present_gl,
                        'dofa' => $dofa,
                        'salary_structure' => $salary_structure,
                        'paypoint' => $paypoint,
                        'bank' => $bank,
                        'acc_no' => $acc_no,
                        // 'dor' => $dor,
                        'old_rank' => $old_rank,
                        'old_gl' => $old_gl,
                        'old_step' => $old_step,
                        'old_salary_per_annum' => $old_salary_per_annum,
                        'new_rank' => $new_rank,
                        'new_gl' => $new_gl,
                        'new_step' => $new_step,
                        'new_salary_per_annum' => $new_salary_per_annum,
                        'effective' => $effective,
                        'placed' => $placed,
                        'months_owed' => $months_owed,
                        'months_paid' => $months_paid,
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'pro_type' => $remark,
                        'remark' => $remark,
                    ]);

                }else if($salary_structure == 'CONHESSHN'){
                    $old_salary_per_annum = Conhesshn::where('conpass_gl', $old_gl)
                                                ->where('conpass_step', $old_step)
                                                ->first()->consolidated_salary_per_annum;

                    $new_salary_per_annum = Conhesshn::where('conpass_gl', $new_gl)
                                                ->where('conpass_step', $new_step)
                                                ->first()->consolidated_salary_per_annum;
                    $variation_amount = $new_salary_per_annum - $old_salary_per_annum;
                    $effective = Carbon::instance($effective);
                    $placed = Carbon::instance($placed);
                    $months_owed = $effective->diffInMonths($placed);
                    $arrears = $variation_amount*$months_owed;
                    $variation = Variation::create([
                        // 'sn' => $sn,
                        'svc_no' => $svc_no,
                        'ippis_no' => $ippis_no,
                        'name' => $name,
                        'gender' => $gender,
                        'dob' => $dob,
                        'present_rank' => $present_rank,
                        'present_gl' => $present_gl,
                        'dofa' => $dofa,
                        'salary_structure' => $salary_structure,
                        'paypoint' => $paypoint,
                        'bank' => $bank,
                        'acc_no' => $acc_no,
                        // 'dor' => $dor,
                        'old_rank' => $old_rank,
                        'old_gl' => $old_gl,
                        'old_step' => $old_step,
                        'old_salary_per_annum' => $old_salary_per_annum,
                        'new_rank' => $new_rank,
                        'new_gl' => $new_gl,
                        'new_step' => $new_step,
                        'new_salary_per_annum' => $new_salary_per_annum,
                        'effective' => $effective,
                        'placed' => $placed,
                        'months_owed' => $months_owed,
                        'months_paid' => $months_paid,
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'pro_type' => $remark,
                        'remark' => $remark,
                    ]);
                }

            });
            
            Alert::success('Variation imported successfully!', 'Success!')->autoclose(222500);
            return back();
        }
    }

    // GENERATE SINGLE ADMIN VARIATION
    public function generate_single_admin_variation($id, DNS2D $dNS2D){   

        $variation = Variation::where('id', $id)->first();
            
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $phpWord->setDefaultFontName('Times New Roman');
        $phpWord->getSettings()->setHideGrammaticalErrors(true);
        $phpWord->getSettings()->setHideSpellingErrors(true);
        $phpWord->setDefaultFontSize(14);

            $current = Carbon::now();
            $currentDate = $current->format('m/Y');
            $progression_code = null;
            if (str_contains($variation->remark, 'UPGRADING/CONVERSION')) { 
                $progression_code = 3037;
            }else{
                $progression_code = 5057;
            }
            
            $image =$dNS2D->getBarcodePNG("
            <h2>".$variation->name."</h2>
            <h3>(".$variation->present_rank.")</h3>
            <h4>Progression:</h4>".'
            <p><b>*</b> '.$variation->old_rank.' <b>to</b> '.$variation->new_rank.' ('.$variation->remark.')</p>', 
            'QRCODE', 10,10);
            $image = str_replace('data:image/png;base64,', '', $image);
            $image = str_replace(' ', '+', $image);
            $imageName = 'QR Code'.'.'.'png';
            File::put(storage_path().'/app/docs/'.$imageName, base64_decode($image));

            // PAGE CONTENT WRAPPER
            $section = $phpWord->addSection([
                'orientation' => 'landscape',
                'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                'marginRight' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                'marginTop' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                'marginBottom' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.38),
                'footerHeight' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.15)
            ]);
            
            // HEADING GOES HERE ////////////////////////////////////////////////
            $section->addText("NIGERIA SECURITY AND CIVIL DEFENCE CORPS ABUJA", ['bold' => true, 'size' => 22], [ 'spaceAfter' => 0, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ]);
            $section->addText("VARIATION ADVICE", ['bold' => true, 'size' => 22], [ 'spaceBefore' => 0,'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ]);


            // TOP TABLE ////////////////////////////////////////////////
            $top_table = $section->addTable(['width' => 100 * 50, 'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT]);
            
            // 1ST ROW
            $top_table->addRow();
            $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(3))
            ->addText("NSCDC Headquarters Office, Abuja", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);

            $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(2.8))
            ->addText("NSCDC/NHQ/$progression_code/$currentDate", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::END ], ['lineHeight' => 2]);
            
            // 2NS ROW
            $top_table->addRow();
            $tag = $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(3))
            ->addTextRun(['lineHeight' => 1, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH ]);
            $tag->addText("FROM:   ", ['size' => 14, 'regular' => true, 'bold' => true]);
            $tag->addText("Personnel Management Department", ['size' => 14, 'regular' => true, 'bold' => false]);

            $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(2.8))
            ->addText("To be submitted not later than the day of the month", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::END ]);
            
            // 3RD ROW
            $top_table->addRow();
            $cell = $top_table->addCell();
            $cell->getStyle()->setGridSpan(2);
            $tag = $cell->addTextRun(['lineHeight' => 1, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH ]);
            $tag->addText("TO:         ", ['size' => 14, 'regular' => true, 'bold' => true]);
            $tag->addText("Variation Control Office/Account Section NSCDC Abuja.", ['size' => 14, 'regular' => true, 'bold' => false]);
            
            // 4TH ROW
            $top_table->addRow();
            $last = $top_table->addCell();
            $last->getStyle()->setGridSpan(2);
            $last->addText("Please find enumerated here under for necessary action a list of Variation for the months of ".$current->format('F, Y'), ['size' => 14, 'regular' => true, 'bold' => false]);
            $section->addTextBreak(1, ['size' => 1]);


            // TABLE BODY OF LETTER //////////////////////////////////////////////////////////////////
            // Table cell style
            $cellStyles = [
                'valign' => 'center'
            ];
            // Table th style
            $thFontStyles = [
                'bold' => true,
                'size' => 12,
            ];
            // Table td style
            $tdFontStyles = [
                'size' => 12,
            ];
            // Table th paragraph style
            $thPstyle = [
                'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER
            ];
            $variation_table = $section->addTable('variation');
            $variation_table->addRow();
            $variation_table->addCell(3000, $cellStyles)->addText('Name', $thFontStyles, $thPstyle);
            $variation_table->addCell(2200, $cellStyles)->addText('Rank/GL', $thFontStyles, $thPstyle);
            $variation_table->addCell(1000, $cellStyles)->addText('Service No.', $thFontStyles, $thPstyle);
            $variation_table->addCell(1500, $cellStyles)->addText('Old Salary Per Annum', $thFontStyles, $thPstyle);
            $variation_table->addCell(1500, $cellStyles)->addText('New Salary Per Annum', $thFontStyles, $thPstyle);
            $variation_table->addCell(null, $cellStyles)->addText('Effective Date', $thFontStyles, $thPstyle);
            $variation_table->addCell(null, $cellStyles)->addText('Variation Amount to', $thFontStyles, $thPstyle);
            $variation_table->addCell(null, $cellStyles)->addText('Remark', $thFontStyles, $thPstyle);
            $variation_table->addCell(null, $cellStyles)->addText('Paypoint', $thFontStyles, $thPstyle);
                
                $variation_table->addRow();
                $variation_table->addCell(3000, $cellStyles)->addText($variation->name, $tdFontStyles);

                $rank = $variation_table->addCell(2200, $cellStyles);
                $rank->addText($variation->new_rank, $tdFontStyles);
                $rank->addText('GL ('.sprintf("%02d", $variation->new_gl).')', $tdFontStyles);

                $variation_table->addCell(1000, $cellStyles)->addText($variation->svc_no, $tdFontStyles);

                $old_salary = $variation_table->addCell(1500, $cellStyles);
                $old_salary->addText('GL ('.sprintf("%02d", $variation->old_gl).'/'.$variation->old_step.')', $tdFontStyles);
                $old_salary->addText('₦'.number_format($variation->old_salary_per_annum), $tdFontStyles);

                $new_salary = $variation_table->addCell(1500, $cellStyles);
                $new_salary->addText('GL ('.sprintf("%02d", $variation->new_gl).'/'.$variation->new_step.')', $tdFontStyles);
                $new_salary->addText('₦'.number_format($variation->new_salary_per_annum), $tdFontStyles);

                $variation_table->addCell(null, $cellStyles)->addText(Carbon::make($variation->effective)->format('d/m/Y'), $tdFontStyles);

                $variation_table->addCell(null, $cellStyles)->addText('₦'.number_format($variation->variation_amount), $tdFontStyles);

                $variation_table->addCell(null, $cellStyles)->addText(ucwords(strtolower($variation->remark)), $tdFontStyles);

                $variation_table->addCell(null, $cellStyles)->addText($variation->paypoint, $tdFontStyles);
            
            $phpWord->addTableStyle(
                'variation', 
                [ 'borderColor' => 'black', 'borderSize'  => 6, 'cellMargin'  => 80], 
                ['bgColor' => 'white', 'bold' => true, 'width' => 100 * 50, 'unit' => 'pct']
            );
            $section->addTextBreak(1, ['size' => 12]);

            // BOTTOM TABLE ////////////////////////////////////////////////
            $footer = $section->addFooter();
            $bottom_table = $footer->addTable(['width' => 100 * 50, 'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT]);
            
            // 1ST ROW
            $bottom_table->addRow();
            $bottom_table->addCell()
            ->addText("SIGNATURE........................................................................................", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);
            $bottom_table->addCell()
            ->addText("DATE RECIEVED.......................................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
            $bottom_table->addCell(2500, ['valign' => 'top', 'vMerge' => 'restart'])
            // QR CODE GOES HERE
            ->addImage(storage_path().'/app/docs/QR Code.png',
            [
                'width' => 95, 'height' => 95, 
                'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END
            ]);
            
            // 2ND ROW
            $bottom_table->addRow();
            $bottom_table->addCell()
            ->addText("OFFICER IN CHARGE VARIATION................................................", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);
            $bottom_table->addCell()
            ->addText("DATE ACTION TAKEN.............................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
            $bottom_table->addCell(2500, ['vMerge' => 'continue']);

            // 3RD ROW
            $bottom_table->addRow();
            $bottom_table->addCell()
            ->addText("FOR: COMMANDANT GENERAL OF NSCDC", ['size' => 16, 'bold' => true, 'italic' => true], ['lineHeight' => 2]);
            $bottom_table->addCell()
            ->addText("SIGNATURE...............................................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
            $bottom_table->addCell(2500, ['vMerge' => 'continue']);

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save(storage_path('app/docs/'.$variation->name.'.docx'));
        return response()->download(storage_path('app/docs/'.$variation->name.'.docx'));
    }
    // GENERATE BULK ADMIN VARIATION
    public function generate_bulk_admin_variation(Request $request, DNS2D $dNS2D){

        $variations = Variation::orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')")->find($request->candidates);

        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $phpWord->setDefaultFontName('Times New Roman');
        $phpWord->getSettings()->setHideGrammaticalErrors(true);
        $phpWord->getSettings()->setHideSpellingErrors(true);
        $phpWord->setDefaultFontSize(14);
        // return count($candidates);
        if(count($variations) > 0){
            foreach($variations as $variation){
        
                $current = Carbon::now();
                $currentDate = $current->format('m/Y');
                $progression_code = null;
                if (str_contains($variation->remark, 'UPGRADING/CONVERSION')) { 
                    $progression_code = 3037;
                }else{
                    $progression_code = 5057;
                }
                
                $image = $dNS2D->getBarcodePNG("
                <h2>".$variation->name."</h2>
                <h3>(".$variation->present_rank.")</h3>
                <h4>Progression:</h4>".'
                <p><b>*</b> '.$variation->old_rank.' <b>to</b> '.$variation->new_rank.' ('.$variation->remark.')</p>', 
                'QRCODE', 10,10);
                $image = str_replace('data:image/png;base64,', '', $image);
                $image = str_replace(' ', '+', $image);
                $imageName = 'QR Code_'.$variation->ippis_no.'_'.$variation->new_gl.'.png';
                File::put(storage_path().'/app/docs/'.$imageName, base64_decode($image));

                // PAGE CONTENT WRAPPER
                $section = $phpWord->addSection([
                    'orientation' => 'landscape',
                    'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                    'marginRight' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                    'marginTop' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.5),
                    'marginBottom' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.38),
                    'footerHeight' => \PhpOffice\PhpWord\Shared\Converter::inchToTwip(0.15)
                ]);
                
                // HEADING GOES HERE ////////////////////////////////////////////////
                $section->addText("NIGERIA SECURITY AND CIVIL DEFENCE CORPS ABUJA", ['bold' => true, 'size' => 22], [ 'spaceAfter' => 0, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ]);
                $section->addText("VARIATION ADVICE", ['bold' => true, 'size' => 22], [ 'spaceBefore' => 0,'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ]);


                // TOP TABLE ////////////////////////////////////////////////
                $top_table = $section->addTable(['width' => 100 * 50, 'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT]);
                
                // 1ST ROW
                $top_table->addRow();
                $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(3))
                ->addText("NSCDC Headquarters Office, Abuja", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);

                $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(2.8))
                ->addText("NSCDC/NHQ/$progression_code/$currentDate", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::END ], ['lineHeight' => 2]);
                
                // 2NS ROW
                $top_table->addRow();
                $tag = $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(3))
                ->addTextRun(['lineHeight' => 1, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH ]);
                $tag->addText("FROM:   ", ['size' => 14, 'regular' => true, 'bold' => true]);
                $tag->addText("Personnel Management Department", ['size' => 14, 'regular' => true, 'bold' => false]);

                $top_table->addCell(\PhpOffice\PhpWord\Shared\Converter::inchToTwip(2.8))
                ->addText("To be submitted not later than the day of the month", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::END ]);
                
                // 3RD ROW
                $top_table->addRow();
                $cell = $top_table->addCell();
                $cell->getStyle()->setGridSpan(2);
                $tag = $cell->addTextRun(['lineHeight' => 1, 'align' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH ]);
                $tag->addText("TO:         ", ['size' => 14, 'regular' => true, 'bold' => true]);
                $tag->addText("Variation Control Office/Account Section NSCDC Abuja.", ['size' => 14, 'regular' => true, 'bold' => false]);
                
                // 4TH ROW
                $top_table->addRow();
                $last = $top_table->addCell();
                $last->getStyle()->setGridSpan(2);
                $last->addText("Please find enumerated here under for necessary action a list of Variation for the months of ".$current->format('F, Y'), ['size' => 14, 'regular' => true, 'bold' => false]);
                $section->addTextBreak(1, ['size' => 1]);


                // TABLE BODY OF LETTER //////////////////////////////////////////////////////////////////
                // Table cell style
                $cellStyles = [
                    'valign' => 'center'
                ];
                // Table th style
                $thFontStyles = [
                    'bold' => true,
                    'size' => 12,
                ];
                // Table td style
                $tdFontStyles = [
                    'size' => 12,
                ];
                // Table th paragraph style
                $thPstyle = [
                    'align' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER
                ];
                $variation_table = $section->addTable('variation');
                $variation_table->addRow();
                $variation_table->addCell(3000, $cellStyles)->addText('Name', $thFontStyles, $thPstyle);
                $variation_table->addCell(2200, $cellStyles)->addText('Rank/GL', $thFontStyles, $thPstyle);
                $variation_table->addCell(1000, $cellStyles)->addText('Service No.', $thFontStyles, $thPstyle);
                $variation_table->addCell(1500, $cellStyles)->addText('Old Salary Per Annum', $thFontStyles, $thPstyle);
                $variation_table->addCell(1500, $cellStyles)->addText('New Salary Per Annum', $thFontStyles, $thPstyle);
                $variation_table->addCell(null, $cellStyles)->addText('Effective Date', $thFontStyles, $thPstyle);
                $variation_table->addCell(null, $cellStyles)->addText('Variation Amount to', $thFontStyles, $thPstyle);
                $variation_table->addCell(null, $cellStyles)->addText('Remark', $thFontStyles, $thPstyle);
                $variation_table->addCell(null, $cellStyles)->addText('Paypoint', $thFontStyles, $thPstyle);
                
                    $variation_table->addRow();
                    $variation_table->addCell(3000, $cellStyles)->addText($variation->name, $tdFontStyles);

                    $rank = $variation_table->addCell(2200, $cellStyles);
                    $rank->addText($variation->new_rank, $tdFontStyles);
                    $rank->addText('GL ('.sprintf("%02d", $variation->new_gl).')', $tdFontStyles);

                    $variation_table->addCell(1000, $cellStyles)->addText($variation->svc_no, $tdFontStyles);

                    $old_salary = $variation_table->addCell(1500, $cellStyles);
                    $old_salary->addText('GL ('.sprintf("%02d", $variation->old_gl).'/'.$variation->old_step.')', $tdFontStyles);
                    $old_salary->addText('₦'.number_format($variation->old_salary_per_annum), $tdFontStyles);

                    $new_salary = $variation_table->addCell(1500, $cellStyles);
                    $new_salary->addText('GL ('.sprintf("%02d", $variation->new_gl).'/'.$variation->new_step.')', $tdFontStyles);
                    $new_salary->addText('₦'.number_format($variation->new_salary_per_annum), $tdFontStyles);

                    $variation_table->addCell(null, $cellStyles)->addText(Carbon::make($variation->effective)->format('d/m/Y'), $tdFontStyles);

                    $variation_table->addCell(null, $cellStyles)->addText('₦'.number_format($variation->variation_amount), $tdFontStyles);

                    $variation_table->addCell(null, $cellStyles)->addText(ucwords(strtolower($variation->remark)), $tdFontStyles);

                    $variation_table->addCell(null, $cellStyles)->addText($variation->paypoint, $tdFontStyles);

                $phpWord->addTableStyle(
                    'variation', 
                    [ 'borderColor' => 'black', 'borderSize'  => 6, 'cellMargin'  => 80], 
                    ['bgColor' => 'white', 'bold' => true, 'width' => 100 * 50, 'unit' => 'pct']
                );
                $section->addTextBreak(1, ['size' => 12]);

                // BOTTOM TABLE ////////////////////////////////////////////////
                $footer = $section->addFooter();
                $bottom_table = $footer->addTable(['width' => 100 * 50, 'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT]);
                
                // 1ST ROW
                $bottom_table->addRow();
                $bottom_table->addCell()
                ->addText("SIGNATURE........................................................................................", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);
                $bottom_table->addCell()
                ->addText("DATE RECIEVED.......................................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
                $bottom_table->addCell(2500, ['valign' => 'top', 'vMerge' => 'restart'])
                // QR CODE GOES HERE
                ->addImage(storage_path().'/app/docs/QR Code_'.$variation->ippis_no.'_'.$variation->new_gl.'.png',
                [
                    'width' => 95, 'height' => 95, 
                    'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::END
                ]);
                
                // 2ND ROW
                $bottom_table->addRow();
                $bottom_table->addCell()
                ->addText("OFFICER IN CHARGE VARIATION................................................", ['size' => 14, 'regular' => true], ['lineHeight' => 2]);
                $bottom_table->addCell()
                ->addText("DATE ACTION TAKEN.............................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
                $bottom_table->addCell(2500, ['vMerge' => 'continue']);

                // 3RD ROW
                $bottom_table->addRow();
                $bottom_table->addCell()
                ->addText("FOR: COMMANDANT GENERAL OF NSCDC", ['size' => 16, 'bold' => true, 'italic' => true], ['lineHeight' => 2]);
                $bottom_table->addCell()
                ->addText("SIGNATURE...............................................", null, [ 'align' => \PhpOffice\PhpWord\SimpleType\Jc::START ], ['lineHeight' => 2]);
                $bottom_table->addCell(2500, ['vMerge' => 'continue']);
    
            }
        
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $objWriter->save(storage_path('app/docs/bulk_variation_advice.docx'));
            return response()->download(storage_path('app/docs/bulk_variation_advice.docx'));

        }else{
            return false;
        }
    }
    
    public function get_months_owed($start_date, $end_date){
        $s = implode("-", array_reverse(explode("/", $start_date)) );
        $e = implode("-", array_reverse(explode("/", $end_date)) );
    
        // get the parts separated
        $start = explode("-",$s);
        $end = explode("-",$e) ;

        $iterations = ((intVal($end[0]) - intVal($start[0])) * 12) - (intVal($start[1]) - intVal($end[1])) ;

        $sets=[$start[0] => array("start" => $s, "end" => "", "months" => 0)];
        $curdstart= $curd = $s;
        $curyear = date("Y", strtotime($s));

        for($x=0; $x<=$iterations; $x++) {
            $curdend = date("Y-m-d", strtotime($curd . " +{$x} months"));
            $curyear = date("Y", strtotime($curdend));
            if (!isset($sets[$curyear])) {
                $sets[$curyear]= array("start" => $curdend, "end" => "", "months" => 0);
            }
            $sets[$curyear]['months']++;
            $sets[$curyear]['end'] = date("Y-m-", strtotime($curdend)) . "31";
            
        }

        return $sets;
    }
    // GENERATE SINGLE FINANCE VARIATION
    public function generate_single_finance_variation($id, DNS2D $dNS2D){   

        $variation = Variation::where('id', $id)->first();
        $effective = Carbon::create($variation->effective);
        $placed = Carbon::create($variation->placed);
        $months_owed = $effective->diffInMonths($placed)+1;

        $iteration = $this->get_months_owed($effective, $placed);
        
        // INSTANTIATE PHP_EXCEL
        $spreadsheet = new Spreadsheet();
        
        // DEFAULT SETUPS
        $spreadsheet->getDefaultStyle()->getFont()->setSize(16);
        $sheet = $spreadsheet->getActiveSheet();

        // PRINT SETUP
        $sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
        $sheet->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
        $sheet->getPageSetup()->setScale(100);
        $sheet->getPageSetup()->setFitToWidth(1); 

        
        $sheet->getSheetView()->setZoomScale(80);
        $sheet->getColumnDimension('A')->setWidth(26);
        $sheet->getColumnDimension('B')->setWidth(30);
        $sheet->getColumnDimension('C')->setAutoSize(true);
        $sheet->getColumnDimension('D')->setAutoSize(true);
        $sheet->getColumnDimension('E')->setAutoSize(true);
        $sheet->getColumnDimension('F')->setAutoSize(true);
        $sheet->getColumnDimension('G')->setAutoSize(true);
        $sheet->getColumnDimension('H')->setAutoSize(true);

        // HEADING STYLE
        $sheet->mergeCells('A1:H1');
        $sheet->getRowDimension(1)->setRowHeight(50);
        $sheet->getStyle('A1:H1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1:H1')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
        $sheet->getStyle('A1:H1')->getFont()->setSize(28);
        $sheet->getStyle('A1:H1')->getFont()->setBold(true);

        $sheet->mergeCells('A2:H2');
        $sheet->getRowDimension(2)->setRowHeight(40);
        $sheet->getStyle('A2:H2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A2:H2')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $sheet->getStyle('A2:H2')->getFont()->setSize(20);
        $sheet->getStyle('A2:H2')->getFont()->setBold(true);

        // HEADING DATA
        $sheet->getCell('A1')->setValue('NIGERIA SECURITY & CIVIL DEFENCE CORPS');
        $sheet->getCell('A2')->setValue('Variation Control');


        // PERSONNEL INFO STYLE
        $sheet->getStyle('A3:B8')->applyFromArray(
            [
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]
        );
        $sheet->getStyle('B3:B8')->getAlignment()->setWrapText(true);
        $sheet->getStyle('A3:A8')->applyFromArray(
            [
                'font' => [
                    'bold' => true,
                ]
            ]
        );
        $sheet->getStyle('A7:H8')->applyFromArray(
            [
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]
        );
        $sheet->mergeCells('C7:H7'); //BANK NAME
        $sheet->mergeCells('B8:H8');
        
        // PERSONNEL INFO DATA
        $sheet->fromArray([
            ['NAME', $variation->name,],
            ['SERVICE NO:-',   $variation->svc_no,],
            ['NHQ FILE NO:-',   '',],
            ['IPPIS NO:-',   $variation->ippis_no,],
            ['ACCOUNT NO/BANK',   $variation->acc_no, $variation->bank],
            ['PURPOSE',   $variation->remark.' ARREARS'],
        ], null, 'A3');

        // LOOP THROUGH THE YEARS
        $i = 1;
        $stp_add = 0;
        $count = 1;
        $loop_count = 0;
        $total = [];
        foreach ($iteration as $key => $value) {
                    
            $old_gl = $variation->old_gl;
            $new_gl = $variation->new_gl;

            $old_step = $variation->old_step+$stp_add;
            $new_step = $variation->new_step+$stp_add;
            
            if($old_gl >= 15 && $old_gl <= 15 && $old_step > 6){
                $old_step = 6;
            }
            elseif($old_gl >= 14 && $old_gl <= 14 && $old_step > 7){
                $old_step = 7;
            }
            elseif($old_gl >= 11 && $old_gl <= 13 && $old_step > 8){
                $old_step = 8;
            }
            elseif($old_gl >= 3 && $old_gl <= 10 && $old_step > 10){
                $old_step = 10;
            }

            if($new_gl >= 16 && $new_gl <= 16 && $new_step > 5){
                $new_step = 5;
            }
            elseif($new_gl >= 15 && $new_gl <= 15 && $new_step > 6){
                $new_step = 6;
            }
            elseif($new_gl >= 14 && $new_gl <= 14 && $new_step > 7){
                $new_step = 7;
            }
            elseif($new_gl >= 11 && $new_gl <= 13 && $new_step > 8){
                $new_step = 8;
            }
            elseif($new_gl >= 3 && $new_gl <= 10 && $new_step > 10){
                $new_step = 10;
            }
                    
            if($variation->salary_structure == 'CONPASS'){
                $old_salary = OldConpass::where('gl', $old_gl)
                            ->where('step', $old_step)
                            ->first();
                $new_salary = OldConpass::where('gl', $new_gl)
                            ->where('step', $new_step)
                            ->first();
            }
            elseif($variation->salary_structure == 'CONMESS'){
                $old_salary = Conmess::where('conpass_gl', $old_gl)
                            ->where('conpass_step', $old_step)
                            ->first();
                $new_salary = Conmess::where('conpass_gl', $new_gl)
                            ->where('conpass_step', $new_step)
                            ->first();
            }
            elseif($variation->salary_structure == 'CONHESSP'){
                $old_salary = Conhessp::where('conpass_gl', $old_gl)
                            ->where('conpass_step', $old_step)
                            ->first();
                $new_salary = Conhessp::where('conpass_gl', $new_gl)
                            ->where('conpass_step', $new_step)
                            ->first();
            }
            elseif($variation->salary_structure == 'CONHESSHN'){
                $old_salary = Conhesshn::where('conpass_gl', $old_gl)
                            ->where('conpass_step', $old_step)
                            ->first();
                $new_salary = Conhesshn::where('conpass_gl', $new_gl)
                            ->where('conpass_step', $new_step)
                            ->first();
            }
            
                            
            array_push($total, ($new_salary->net_pay - $old_salary->net_pay) * ($value['months'] - $variation->months_paid));

            // MAIN TABLE STYLES
            $sheet->mergeCells('A'.($i+8).':H'.($i+8));
            $sheet->getStyle('A'.($i+8).':H'.($i+8))->applyFromArray( //YEAR HEADING
                [
                    'font' => [
                        'bold' => true,
                    ],
                    'alignment' => [
                        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                    ],
                ]
            );
            $sheet->getStyle('A'.($i+8).':H'.($i+14))->applyFromArray([
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                ],
            ]);
            $sheet->getStyle('A'.($i+8))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                    $sheet->getRowDimension($i+8)->setRowHeight(40);
            $sheet->getStyle('A'.($i+9).':H'.($i+9))->applyFromArray( //COLUMN HEADS
                [
                    'font' => [
                        'bold' => true,
                    ],
                    'alignment' => [
                        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                    ],
                ]
            );
            $sheet->getStyle('C'.($i+10).':H'.($i+14))->getNumberFormat() //COMPUTATIONS
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $sheet->getStyle('A'.($i+13).':H'.($i+14))->getAlignment()->setWrapText(true); //TOTALS
            $sheet->getStyle('A'.($i+13).':H'.($i+14))->applyFromArray( //TOTALS
                [
                    'font' => [
                        'bold' => true,
                    ]
                ]
            );

            // MAIN TABLE DATA
            if ($loop_count == 0) { // FIRST ITERATION
                if ($variation->months_paid > 0) {
                    $sheet->fromArray([
                        ['YEAR '.$key],
                        [
                            date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "m/d/Y"),
                            'GRADE LEVEL',
                            'GROSS',
                            'TAX',
                            'NHF',
                            'PENSION',
                            'TOTAL DED',
                            'NET PAY'
                        ],
                        [
                            'PROMOTED SALARY (GROSS)',
                            $variation->new_gl.'/'.($new_step),
                            $new_salary->gross_emolument,
                            $new_salary->tax,
                            $new_salary->nhf,
                            $new_salary->pension,
                            $new_salary->total_deduction,
                            $new_salary->net_pay
                        ],
                        [
                            'OLD SALARY (GROSS)',
                            $variation->old_gl.'/'.($old_step),
                            $old_salary->gross_emolument,
                            $old_salary->tax,
                            $old_salary->nhf,
                            $old_salary->pension,
                            $old_salary->total_deduction,
                            $old_salary->net_pay
                        ],
                        [
                            '1 MONTH VARIANCE',
                            '',
                            $new_salary->gross_emolument-$old_salary->gross_emolument,
                            $new_salary->tax-$old_salary->tax,
                            $new_salary->nhf-$old_salary->nhf,
                            $new_salary->pension-$old_salary->pension,
                            $new_salary->total_deduction-$old_salary->total_deduction,
                            $new_salary->net_pay-$old_salary->net_pay,
                        ],
                        [
                            $value['months'] - $variation->months_paid.' MONTHS VARIANCE',
                            '',
                            ($new_salary->gross_emolument-$old_salary->gross_emolument)*($value['months'] - $variation->months_paid),
                            ($new_salary->tax-$old_salary->tax)*($value['months'] - $variation->months_paid),
                            ($new_salary->nhf-$old_salary->nhf)*($value['months'] - $variation->months_paid),
                            ($new_salary->pension-$old_salary->pension)*($value['months'] - $variation->months_paid),
                            ($new_salary->total_deduction-$old_salary->total_deduction)*($value['months'] - $variation->months_paid),
                            ($new_salary->net_pay-$old_salary->net_pay)*($value['months'] - $variation->months_paid),
                        ],
                    ], null, "A".($i+8));
                }else{
                    $sheet->fromArray([
                        ['YEAR '.$key],
                        [
                            date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "d/m/Y"),
                            'GRADE LEVEL',
                            'GROSS',
                            'TAX',
                            'NHF',
                            'PENSION',
                            'TOTAL DED',
                            'NET PAY'
                        ],
                        [
                            'PROMOTED SALARY (GROSS)',
                            $variation->new_gl.'/'.($new_step),
                            $new_salary->gross_emolument,
                            $new_salary->tax,
                            $new_salary->nhf,
                            $new_salary->pension,
                            $new_salary->total_deduction,
                            $new_salary->net_pay
                        ],
                        [
                            'OLD SALARY (GROSS)',
                            $variation->old_gl.'/'.($old_step),
                            $old_salary->gross_emolument,
                            $old_salary->tax,
                            $old_salary->nhf,
                            $old_salary->pension,
                            $old_salary->total_deduction,
                            $old_salary->net_pay
                        ],
                        [
                            '1 MONTH VARIANCE',
                            '',
                            $new_salary->gross_emolument-$old_salary->gross_emolument,
                            $new_salary->tax-$old_salary->tax,
                            $new_salary->nhf-$old_salary->nhf,
                            $new_salary->pension-$old_salary->pension,
                            $new_salary->total_deduction-$old_salary->total_deduction,
                            $new_salary->net_pay-$old_salary->net_pay,
                        ],
                        [
                            $value['months'].' MONTHS VARIANCE',
                            '',
                            ($new_salary->gross_emolument-$old_salary->gross_emolument)*($value['months'] - $variation->months_paid),
                            ($new_salary->tax-$old_salary->tax)*($value['months'] - $variation->months_paid),
                            ($new_salary->nhf-$old_salary->nhf)*($value['months'] - $variation->months_paid),
                            ($new_salary->pension-$old_salary->pension)*($value['months'] - $variation->months_paid),
                            ($new_salary->total_deduction-$old_salary->total_deduction)*($value['months'] - $variation->months_paid),
                            ($new_salary->net_pay-$old_salary->net_pay)*($value['months'] - $variation->months_paid),
                        ],
                    ], null, "A".($i+8));
                }
            }else{ //EVERY OTHER ITERATION
                $sheet->fromArray([
                    ['YEAR '.$key],
                    [
                        date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "d/m/Y"),
                        'GRADE LEVEL',
                        'GROSS',
                        'TAX',
                        'NHF',
                        'PENSION',
                        'TOTAL DED',
                        'NET PAY'
                    ],
                    [
                        'PROMOTED SALARY (GROSS)',
                        $variation->new_gl.'/'.($new_step),
                        $new_salary->gross_emolument,
                        $new_salary->tax,
                        $new_salary->nhf,
                        $new_salary->pension,
                        $new_salary->total_deduction,
                        $new_salary->net_pay
                    ],
                    [
                        'OLD SALARY (GROSS)',
                        $variation->old_gl.'/'.($old_step),
                        $old_salary->gross_emolument,
                        $old_salary->tax,
                        $old_salary->nhf,
                        $old_salary->pension,
                        $old_salary->total_deduction,
                        $old_salary->net_pay
                    ],
                    [
                        '1 MONTH VARIANCE',
                        '',
                        $new_salary->gross_emolument-$old_salary->gross_emolument,
                        $new_salary->tax-$old_salary->tax,
                        $new_salary->nhf-$old_salary->nhf,
                        $new_salary->pension-$old_salary->pension,
                        $new_salary->total_deduction-$old_salary->total_deduction,
                        $new_salary->net_pay-$old_salary->net_pay,
                    ],
                    [
                        $value['months'].' MONTHS VARIANCE',
                        '',
                        ($new_salary->gross_emolument-$old_salary->gross_emolument)*$value['months'],
                        ($new_salary->tax-$old_salary->tax)*$value['months'],
                        ($new_salary->nhf-$old_salary->nhf)*$value['months'],
                        ($new_salary->pension-$old_salary->pension)*$value['months'],
                        ($new_salary->total_deduction-$old_salary->total_deduction)*$value['months'],
                        ($new_salary->net_pay-$old_salary->net_pay)*$value['months'],
                    ],
                ], null, "A".($i+8));
            }

            // LAST ITERATION
            if($count == count($iteration)){
                if($variation->months_paid > 0){
                    
                    $sheet->fromArray([
                        [
                            $months_owed - $variation->months_paid.' TOTAL MONTHS VARIANCE',
                            '',
                            '',
                            '',
                            '',
                            '',
                            '',
                            array_sum($total),
                        ],
                    ], null, "A".($i+14));

                     // SIGNATURE AREA
                     $sheet->getStyle("A".($i+18))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                     $sheet->getStyle("A".($i+18))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                     $sheet->getStyle("A".($i+18))->applyFromArray([
                         'borders' => [
                             'top' => [
                                 'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
                                 'color' => ['argb' => '000000'],
                             ],
                         ],
                     ]);
                     $sheet->getRowDimension($i+18)->setRowHeight(25);
                     $sheet->getStyle("A".($i+18))->getFont()->setSize(20);
                     $sheet->getStyle("A".($i+18))->getFont()->setBold(true);
                     $sheet->fromArray([
                         ['OC VARIATION'],
                     ], null, "A".($i+18));

                }else{
                    $sheet->fromArray([
                        [
                            $months_owed.' TOTAL MONTHS VARIANCE',
                            '',
                            '',
                            '',
                            '',
                            '',
                            '',
                            array_sum($total),
                        ],
                    ], null, "A".($i+14));

                     // SIGNATURE AREA
                     $sheet->getStyle("A".($i+18))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                     $sheet->getStyle("A".($i+18))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                     $sheet->getStyle("A".($i+18))->applyFromArray([
                         'borders' => [
                             'top' => [
                                 'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
                                 'color' => ['argb' => '000000'],
                             ],
                         ],
                     ]);
                     $sheet->getRowDimension($i+18)->setRowHeight(25);
                     $sheet->getStyle("A".($i+18))->getFont()->setSize(20);
                     $sheet->getStyle("A".($i+18))->getFont()->setBold(true);
                     $sheet->fromArray([
                         ['OC VARIATION'],
                     ], null, "A".($i+18));
                }
            }

            $i+=6;
            $stp_add++;
            $count++;
            $loop_count++;
        }
        $writer = new Xlsx($spreadsheet);
        $writer->save(storage_path('app/docs/'.$variation->name.'.xlsx'));
        return response()->download(storage_path('app/docs/'.$variation->name.'.xlsx'));
    }
    // GENERATE SINGLE FINANCE VARIATION
    public function generate_bulk_finance_variation(Request $request, DNS2D $dNS2D){   

        $variations = Variation::orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')")->find($request->candidates);
        
        // INSTANTIATION OF PHPEXCEL
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getDefaultStyle()->getFont()->setSize(16);
        $sheet = $spreadsheet->getActiveSheet();

        if (count($variations) > 0) {
            
            foreach ($variations as $sheet_key => $variation) {
                
                $effective = Carbon::create($variation->effective);
                $placed = Carbon::create($variation->placed);
                $months_owed = $effective->diffInMonths($placed)+1;
                $iteration = $this->get_months_owed($effective, $placed);

                $sheet_key > 0 ? $sheet = $spreadsheet->createSheet($sheet_key) : null;
                $sheet->setTitle($variation->ippis_no.' - '.substr($variation->remark, 0, 5));

                // PRINT SETUP
                $sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
                $sheet->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
                $sheet->getPageSetup()->setScale(100);
                $sheet->getPageSetup()->setFitToWidth(1); 

                
                $sheet->getSheetView()->setZoomScale(80);
                $sheet->getColumnDimension('A')->setWidth(26);
                $sheet->getColumnDimension('B')->setWidth(30);
                $sheet->getColumnDimension('C')->setAutoSize(true);
                $sheet->getColumnDimension('D')->setAutoSize(true);
                $sheet->getColumnDimension('E')->setAutoSize(true);
                $sheet->getColumnDimension('F')->setAutoSize(true);
                $sheet->getColumnDimension('G')->setAutoSize(true);
                $sheet->getColumnDimension('H')->setAutoSize(true);

                // HEADING STYLE
                $sheet->mergeCells('A1:H1');
                $sheet->getRowDimension(1)->setRowHeight(50);
                $sheet->getStyle('A1:H1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A1:H1')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                $sheet->getStyle('A1:H1')->getFont()->setSize(28);
                $sheet->getStyle('A1:H1')->getFont()->setBold(true);

                $sheet->mergeCells('A2:H2');
                $sheet->getRowDimension(2)->setRowHeight(40);
                $sheet->getStyle('A2:H2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A2:H2')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
                $sheet->getStyle('A2:H2')->getFont()->setSize(20);
                $sheet->getStyle('A2:H2')->getFont()->setBold(true);

                // HEADING DATA
                $sheet->getCell('A1')->setValue('NIGERIA SECURITY & CIVIL DEFENCE CORPS');
                $sheet->getCell('A2')->setValue('Variation Control');


                // PERSONNEL INFO STYLE
                $sheet->getStyle('A3:B8')->applyFromArray(
                    [
                        'alignment' => [
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                        ],
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                                'color' => ['argb' => '000000'],
                            ],
                        ],
                    ]
                );
                $sheet->getStyle('B3:B8')->getAlignment()->setWrapText(true);
                $sheet->getStyle('A3:A8')->applyFromArray(
                    [
                        'font' => [
                            'bold' => true,
                        ]
                    ]
                );
                $sheet->getStyle('A7:H8')->applyFromArray(
                    [
                        'alignment' => [
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                        ],
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                                'color' => ['argb' => '000000'],
                            ],
                        ],
                    ]
                );
                $sheet->mergeCells('C7:H7'); //BANK NAME
                $sheet->mergeCells('B8:H8');
                
                // PERSONNEL INFO DATA
                $sheet->fromArray([
                    ['NAME', $variation->name,],
                    ['SERVICE NO:-',   $variation->svc_no,],
                    ['NHQ FILE NO:-',   '',],
                    ['IPPIS NO:-',   $variation->ippis_no,],
                    ['ACCOUNT NO/BANK',   $variation->acc_no, $variation->bank],
                    ['PURPOSE',   $variation->remark.' ARREARS'],
                ], null, 'A3');

                // LOOP THROUGH THE YEARS
                $i = 1;
                $stp_add = 0;
                $count = 1;
                $loop_count = 0;
                $total = [];
                foreach ($iteration as $key => $value) {
                            
                    $old_gl = $variation->old_gl;
                    $new_gl = $variation->new_gl;

                    $old_step = $variation->old_step+$stp_add;
                    $new_step = $variation->new_step+$stp_add;
                    
                    if($old_gl >= 15 && $old_gl <= 15 && $old_step > 6){
                        $old_step = 6;
                    }
                    elseif($old_gl >= 14 && $old_gl <= 14 && $old_step > 7){
                        $old_step = 7;
                    }
                    elseif($old_gl >= 11 && $old_gl <= 13 && $old_step > 8){
                        $old_step = 8;
                    }
                    elseif($old_gl >= 3 && $old_gl <= 10 && $old_step > 10){
                        $old_step = 10;
                    }

                    if($new_gl >= 16 && $new_gl <= 16 && $new_step > 5){
                        $new_step = 5;
                    }
                    elseif($new_gl >= 15 && $new_gl <= 15 && $new_step > 6){
                        $new_step = 6;
                    }
                    elseif($new_gl >= 14 && $new_gl <= 14 && $new_step > 7){
                        $new_step = 7;
                    }
                    elseif($new_gl >= 11 && $new_gl <= 13 && $new_step > 8){
                        $new_step = 8;
                    }
                    elseif($new_gl >= 3 && $new_gl <= 10 && $new_step > 10){
                        $new_step = 10;
                    }
                            
                    if($variation->salary_structure == 'CONPASS'){
                        $old_salary = OldConpass::where('gl', $old_gl)
                                    ->where('step', $old_step)
                                    ->first();
                        $new_salary = OldConpass::where('gl', $new_gl)
                                    ->where('step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONMESS'){
                        $old_salary = Conmess::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conmess::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONHESSP'){
                        $old_salary = Conhessp::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conhessp::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONHESSHN'){
                        $old_salary = Conhesshn::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conhesshn::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                                    
                    // return $new_salary;
                    array_push($total, ($new_salary->net_pay - $old_salary->net_pay) * ($value['months'] - $variation->months_paid));

                    // MAIN TABLE STYLES
                    $sheet->mergeCells('A'.($i+8).':H'.($i+8));
                    $sheet->getStyle('A'.($i+8).':H'.($i+8))->getFont()->setBold(true);
                    $sheet->getStyle('A'.($i+8).':H'.($i+14))->applyFromArray([
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                                'color' => ['argb' => '000000'],
                            ],
                        ],
                        'alignment' => [
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                        ],
                    ]);
                    $sheet->getStyle('A'.($i+8))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                    $sheet->getRowDimension($i+8)->setRowHeight(40);

                    $sheet->getStyle('A'.($i+9).':H'.($i+9))->applyFromArray( //COLUMN HEADS
                        [
                            'font' => [
                                'bold' => true,
                            ],
                            'alignment' => [
                                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                            ],
                        ]
                    );
                    $sheet->getStyle('C'.($i+10).':H'.($i+14))->getNumberFormat() //COMPUTATIONS
                    ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $sheet->getStyle('A'.($i+13).':H'.($i+14))->getAlignment()->setWrapText(true); //TOTALS
                    $sheet->getStyle('A'.($i+13).':H'.($i+14))->getFont()->setBold(true);

                    // MAIN TABLE DATA
                    if ($loop_count == 0) { // FIRST ITERATION
                        if ($variation->months_paid > 0) {
                            $sheet->fromArray([
                                ['YEAR '.$key],
                                [
                                    date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "m/d/Y"),
                                    'GRADE LEVEL',
                                    'GROSS',
                                    'TAX',
                                    'NHF',
                                    'PENSION',
                                    'TOTAL DED',
                                    'NET PAY'
                                ],
                                [
                                    'PROMOTED SALARY (GROSS)',
                                    $variation->new_gl.'/'.($new_step),
                                    $new_salary->gross_emolument,
                                    $new_salary->tax,
                                    $new_salary->nhf,
                                    $new_salary->pension,
                                    $new_salary->total_deduction,
                                    $new_salary->net_pay
                                ],
                                [
                                    'OLD SALARY (GROSS)',
                                    $variation->old_gl.'/'.($old_step),
                                    $old_salary->gross_emolument,
                                    $old_salary->tax,
                                    $old_salary->nhf,
                                    $old_salary->pension,
                                    $old_salary->total_deduction,
                                    $old_salary->net_pay
                                ],
                                [
                                    '1 MONTH VARIANCE',
                                    '',
                                    $new_salary->gross_emolument-$old_salary->gross_emolument,
                                    $new_salary->tax-$old_salary->tax,
                                    $new_salary->nhf-$old_salary->nhf,
                                    $new_salary->pension-$old_salary->pension,
                                    $new_salary->total_deduction-$old_salary->total_deduction,
                                    $new_salary->net_pay-$old_salary->net_pay,
                                ],
                                [
                                    $value['months'] - $variation->months_paid.' MONTHS VARIANCE',
                                    '',
                                    ($new_salary->gross_emolument-$old_salary->gross_emolument)*($value['months'] - $variation->months_paid),
                                    ($new_salary->tax-$old_salary->tax)*($value['months'] - $variation->months_paid),
                                    ($new_salary->nhf-$old_salary->nhf)*($value['months'] - $variation->months_paid),
                                    ($new_salary->pension-$old_salary->pension)*($value['months'] - $variation->months_paid),
                                    ($new_salary->total_deduction-$old_salary->total_deduction)*($value['months'] - $variation->months_paid),
                                    ($new_salary->net_pay-$old_salary->net_pay)*($value['months'] - $variation->months_paid),
                                ],
                            ], null, "A".($i+8));
                        }else{
                            $sheet->fromArray([
                                ['YEAR '.$key],
                                [
                                    date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "d/m/Y"),
                                    'GRADE LEVEL',
                                    'GROSS',
                                    'TAX',
                                    'NHF',
                                    'PENSION',
                                    'TOTAL DED',
                                    'NET PAY'
                                ],
                                [
                                    'PROMOTED SALARY (GROSS)',
                                    $variation->new_gl.'/'.($new_step),
                                    $new_salary->gross_emolument,
                                    $new_salary->tax,
                                    $new_salary->nhf,
                                    $new_salary->pension,
                                    $new_salary->total_deduction,
                                    $new_salary->net_pay
                                ],
                                [
                                    'OLD SALARY (GROSS)',
                                    $variation->old_gl.'/'.($old_step),
                                    $old_salary->gross_emolument,
                                    $old_salary->tax,
                                    $old_salary->nhf,
                                    $old_salary->pension,
                                    $old_salary->total_deduction,
                                    $old_salary->net_pay
                                ],
                                [
                                    '1 MONTH VARIANCE',
                                    '',
                                    $new_salary->gross_emolument-$old_salary->gross_emolument,
                                    $new_salary->tax-$old_salary->tax,
                                    $new_salary->nhf-$old_salary->nhf,
                                    $new_salary->pension-$old_salary->pension,
                                    $new_salary->total_deduction-$old_salary->total_deduction,
                                    $new_salary->net_pay-$old_salary->net_pay,
                                ],
                                [
                                    $value['months'].' MONTHS VARIANCE',
                                    '',
                                    ($new_salary->gross_emolument-$old_salary->gross_emolument)*($value['months'] - $variation->months_paid),
                                    ($new_salary->tax-$old_salary->tax)*($value['months'] - $variation->months_paid),
                                    ($new_salary->nhf-$old_salary->nhf)*($value['months'] - $variation->months_paid),
                                    ($new_salary->pension-$old_salary->pension)*($value['months'] - $variation->months_paid),
                                    ($new_salary->total_deduction-$old_salary->total_deduction)*($value['months'] - $variation->months_paid),
                                    ($new_salary->net_pay-$old_salary->net_pay)*($value['months'] - $variation->months_paid),
                                ],
                            ], null, "A".($i+8));
                        }
                    }else{ //EVERY OTHER ITERATION
                        $sheet->fromArray([
                            ['YEAR '.$key],
                            [
                                date_format(date_create($value['start']), "d/m/Y").' - '.date_format(date_create($value['end']), "d/m/Y"),
                                'GRADE LEVEL',
                                'GROSS',
                                'TAX',
                                'NHF',
                                'PENSION',
                                'TOTAL DED',
                                'NET PAY'
                            ],
                            [
                                'PROMOTED SALARY (GROSS)',
                                $variation->new_gl.'/'.($new_step),
                                $new_salary->gross_emolument,
                                $new_salary->tax,
                                $new_salary->nhf,
                                $new_salary->pension,
                                $new_salary->total_deduction,
                                $new_salary->net_pay
                            ],
                            [
                                'OLD SALARY (GROSS)',
                                $variation->old_gl.'/'.($old_step),
                                $old_salary->gross_emolument,
                                $old_salary->tax,
                                $old_salary->nhf,
                                $old_salary->pension,
                                $old_salary->total_deduction,
                                $old_salary->net_pay
                            ],
                            [
                                '1 MONTH VARIANCE',
                                '',
                                $new_salary->gross_emolument-$old_salary->gross_emolument,
                                $new_salary->tax-$old_salary->tax,
                                $new_salary->nhf-$old_salary->nhf,
                                $new_salary->pension-$old_salary->pension,
                                $new_salary->total_deduction-$old_salary->total_deduction,
                                $new_salary->net_pay-$old_salary->net_pay,
                            ],
                            [
                                $value['months'].' MONTHS VARIANCE',
                                '',
                                ($new_salary->gross_emolument-$old_salary->gross_emolument)*$value['months'],
                                ($new_salary->tax-$old_salary->tax)*$value['months'],
                                ($new_salary->nhf-$old_salary->nhf)*$value['months'],
                                ($new_salary->pension-$old_salary->pension)*$value['months'],
                                ($new_salary->total_deduction-$old_salary->total_deduction)*$value['months'],
                                ($new_salary->net_pay-$old_salary->net_pay)*$value['months'],
                            ],
                        ], null, "A".($i+8));
                    }

                    // LAST ITERATION
                    if($count == count($iteration)){
                        if($variation->months_paid > 0){
                            
                            $sheet->fromArray([
                                [
                                    $months_owed - $variation->months_paid.' TOTAL MONTHS VARIANCE',
                                    '',
                                    '',
                                    '',
                                    '',
                                    '',
                                    '',
                                    array_sum($total),
                                ],
                            ], null, "A".($i+14));

                            // SIGNATURE AREA
                            $sheet->getStyle("A".($i+18))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                            $sheet->getStyle("A".($i+18))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                            $sheet->getStyle("A".($i+18))->applyFromArray([
                                'borders' => [
                                    'top' => [
                                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
                                        'color' => ['argb' => '000000'],
                                    ],
                                ],
                            ]);
                            $sheet->getRowDimension($i+18)->setRowHeight(25);
                            $sheet->getStyle("A".($i+18))->getFont()->setSize(20);
                            $sheet->getStyle("A".($i+18))->getFont()->setBold(true);
                            $sheet->fromArray([
                                ['OC VARIATION'],
                            ], null, "A".($i+18));
                            
                        }else{
                            $sheet->fromArray([
                                [
                                    $months_owed.' TOTAL MONTHS VARIANCE',
                                    '',
                                    '',
                                    '',
                                    '',
                                    '',
                                    '',
                                    array_sum($total),
                                ],
                            ], null, "A".($i+14));

                            // SIGNATURE AREA
                            $sheet->getStyle("A".($i+18))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                            $sheet->getStyle("A".($i+18))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                            $sheet->getStyle("A".($i+18))->applyFromArray([
                                'borders' => [
                                    'top' => [
                                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
                                        'color' => ['argb' => '000000'],
                                    ],
                                ],
                            ]);
                            $sheet->getRowDimension($i+18)->setRowHeight(25);
                            $sheet->getStyle("A".($i+18))->getFont()->setSize(20);
                            $sheet->getStyle("A".($i+18))->getFont()->setBold(true);
                            $sheet->fromArray([
                                ['OC VARIATION'],
                            ], null, "A".($i+18));
                        }
                    }
                    $i+=6;
                    $stp_add++;
                    $count++;
                    $loop_count++;
                }
                $sheet_key++;
            }

            $writer = new Xlsx($spreadsheet);
            $writer->save(storage_path('app/excel/bulk_variation_advice.xlsx'));
            return response()->download(storage_path('app/excel/bulk_variation_advice.xlsx'));
        }
        else{
            return false;
        }
    }

    // GENERATE SINGLE IPPIS TRANSLATION
    public function generate_bulk_ippis_translation(Request $request, DNS2D $dNS2D){   

        $variations = Variation::orderByRaw("FIELD(present_rank, 'CG', 'DCG', 'ACG', 'CC', 'DCC', 'ACC', 'CSC', 'SC', 'DSC', 'ASC I', 'ASC II', 'CIC', 'DCIC', 'ACIC', 'PIC I', 'PIC II', 'SIC', 'IC', 'AIC', 'CCA', 'SCA', 'CA I', 'CA II', 'CA III', 'NON UNIFORM', 'N/A')")->find($request->candidates);
        
        // DEFAULT SETUPS
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getDefaultStyle()->getFont()->setSize(16);
        $sheet = $spreadsheet->getActiveSheet();

        if (count($variations) > 0) {
            
            // PRINT SETUP
            $sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
            $sheet->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
            // $sheet->getPageSetup()->setScale(100);
            $sheet->getSheetView()->setZoomScale(80);
            $sheet->getColumnDimension('A')->setAutoSize(true);
            $sheet->getColumnDimension('B')->setAutoSize(true);
            $sheet->getColumnDimension('C')->setAutoSize(true);
            $sheet->getColumnDimension('D')->setAutoSize(true);
            $sheet->getColumnDimension('E')->setAutoSize(true);
            $sheet->getColumnDimension('F')->setAutoSize(true);
            $sheet->getColumnDimension('G')->setAutoSize(true);
            $sheet->getColumnDimension('H')->setAutoSize(true);
            $sheet->getColumnDimension('I')->setAutoSize(true);
            $sheet->getColumnDimension('J')->setAutoSize(true);
            $sheet->getColumnDimension('K')->setAutoSize(true);
            $sheet->getColumnDimension('L')->setAutoSize(true);
            $sheet->getColumnDimension('M')->setAutoSize(true);
            $sheet->getColumnDimension('N')->setAutoSize(true);
            $sheet->getColumnDimension('O')->setAutoSize(true);

            // ENTER THE HEADING
            $sheet->getStyle('A1:O1')->getFont()->setBold(true);
            $sheet->fromArray([
                ['S/N', 'NAME', 'SVC NO', 'IPPIS NO.', 'ACC NO.', 'OLD GL', 'STEP', 'NEW GL', 'STEP', 'EFFECTIVE DATE', 'PLACEMENT DATE', 'NO. OF MONTHS', 'MONTHS PAID', 'ARREARS', 'REMARK']
            ]);

            $count = 2;
            $sn = 1;
            foreach ($variations as $main_key => $variation) {
                
                $effective = Carbon::create($variation->effective);
                $placed = Carbon::create($variation->placed);
                $iteration = $this->get_months_owed($effective, $placed);

                // LOOP THROUGH THE YEARS
                $stp_add = 0;
                $loop_count = 0;
                $total_net = [];
                $total_month = [];

                foreach ($iteration as $key => $value) {

                    $old_gl = $variation->old_gl;
                    $new_gl = $variation->new_gl;

                    $old_step = $variation->old_step+$stp_add;
                    $new_step = $variation->new_step+$stp_add;
                    
                    if ($old_gl >= 15 && $old_gl <= 15 && $old_step > 6) {
                        $old_step = 6;
                    } elseif ($old_gl >= 14 && $old_gl <= 14 && $old_step > 7) {
                        $old_step = 7;
                    } elseif ($old_gl >= 11 && $old_gl <= 13 && $old_step > 8) {
                        $old_step = 8;
                    } elseif ($old_gl >= 3 && $old_gl <= 10 && $old_step > 10) {
                        $old_step = 10;
                    }

                    if ($new_gl >= 16 && $new_gl <= 16 && $new_step > 5) {
                        $new_step = 5;
                    } elseif ($new_gl >= 15 && $new_gl <= 15 && $new_step > 6) {
                        $new_step = 6;
                    } elseif ($new_gl >= 14 && $new_gl <= 14 && $new_step > 7) {
                        $new_step = 7;
                    } elseif ($new_gl >= 11 && $new_gl <= 13 && $new_step > 8) {
                        $new_step = 8;
                    } elseif ($new_gl >= 3 && $new_gl <= 10 && $new_step > 10) {
                        $new_step = 10;
                    }
                            
                    if($variation->salary_structure == 'CONPASS'){
                        $old_salary = OldConpass::where('gl', $old_gl)
                                    ->where('step', $old_step)
                                    ->first();
                        $new_salary = OldConpass::where('gl', $new_gl)
                                    ->where('step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONMESS'){
                        $old_salary = Conmess::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conmess::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONHESSP'){
                        $old_salary = Conhessp::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conhessp::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                    elseif($variation->salary_structure == 'CONHESSHN'){
                        $old_salary = Conhesshn::where('conpass_gl', $old_gl)
                                    ->where('conpass_step', $old_step)
                                    ->first();
                        $new_salary = Conhesshn::where('conpass_gl', $new_gl)
                                    ->where('conpass_step', $new_step)
                                    ->first();
                    }
                    
                    $months = $value['months'];
                    $months_paid = $variation->months_paid;
                    $new_net_pay = $new_salary->net_pay;
                    $old_net_pay = $old_salary->net_pay;
                    $net = 0;
                    $no_months = 0;
                    if ($loop_count == 0) {

                        if ($months_paid > 0) {
                            $net = ($new_net_pay - $old_net_pay) * ($months - $months_paid);
                            $no_months = $months;
                        } else {
                            $net = ($new_net_pay - $old_net_pay) * $months;
                            $no_months = $months;
                        }

                    } else {

                        $net = ($new_net_pay - $old_net_pay) * $months;
                        $no_months = $months;
                    }

                    array_push($total_month, $no_months);
                    array_push($total_net, $net);

                    $stp_add++;
                    $loop_count++;
                }
                
                $sheet->getStyle('E'.($main_key+1))->getNumberFormat()->setFormatCode('0000000000');
                $sheet->getStyle('N'.($main_key+1))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                $sheet->getStyle('J'.($main_key+1))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DMYSLASH);
                $sheet->getStyle('K'.($main_key+1))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DMYSLASH);

                $sheet->fromArray([
                    [
                        $main_key+1,
                        $variation->name,
                        $variation->svc_no,
                        $variation->ippis_no,
                        $variation->acc_no,
                        $variation->old_gl,
                        $variation->old_step,
                        $variation->new_gl,
                        $variation->new_step,
                        Carbon::create($variation->effective)->format('d/m/Y'),
                        Carbon::create($variation->placed)->format('d/m/Y'),
                        array_sum($total_month),
                        $variation->months_paid,
                        array_sum($total_net),
                        $variation->remark
                    ],
                   
                ], null, "A".($count));

                // LAST ITERATION
                if($sn == count($variations)){

                    // ALIGN WHOLE TABLE LEFT
                    $sheet->getStyle('A2:O'.($count+1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                    // FORMAT TOTAL CELL
                    $sheet->getStyle('N2:N'.($count+1))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);

                    // ADD A TOTAL
                    $sheet->setCellValue("N".($count+1),'=SUM(N2:N'.$count.')');
                            
                }
                
                $sn++;
                $count++;
            }

            $writer = new Xlsx($spreadsheet);
            $writer->save(storage_path('app/excel/ippis_translation.xlsx'));
            return response()->download(storage_path('app/excel/ippis_translation.xlsx'));
        }
        else{
            return false;
        }
    }

}