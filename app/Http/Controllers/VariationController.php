<?php

namespace App\Http\Controllers;

use App\Models\Variation;
use App\Models\Beneficiary;
use App\Models\Conhesshn;
use App\Models\Conhessp;
use App\Models\Conmess;
use App\Models\Conpass;
use App\Models\OldConpass;
use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\File;
use Milon\Barcode\DNS2D;
use Rap2hpoutre\FastExcel\FastExcel;
use RealRashid\SweetAlert\Facades\Alert;
use Yajra\DataTables\Facades\DataTables;

class VariationController extends Controller
{
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
                    <a href="'.route('generate_single_variation_slip', $variation->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
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
                    <a href="'.route('generate_single_variation_slip', $variation->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
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
                    <a href="'.route('generate_single_variation_slip', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
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
                    <a href="'.route('generate_single_variation_slip', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
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
                    <a href="'.route('generate_single_variation_slip', $variarion->id).'" style="margin-right:5px;" class="blue-text text-darken-3" title="Print variation slip"><i class="fas fa-file-word fa-lg"></i></a>
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
                $line['dor'] == '' ? $dor = null : $dor = $line['dor'];
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
                    'dor' => $dor,
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
                $line['dor'] == '' ? $dor = null : $dor = $line['dor'];

                $line['old_rank'] == '' ? $old_rank = null : $old_rank = $line['old_rank'];
                $line['old_gl'] == '' ? $old_gl = null : $old_gl = $line['old_gl'];
                $line['old_step'] == '' ? $old_step = null : $old_step = $line['old_step'];
            
                $line['new_rank'] == '' ? $new_rank = null : $new_rank = $line['new_rank'];
                $line['new_gl'] == '' ? $new_gl = null : $new_gl = $line['new_gl'];
                $line['new_step'] == '' ? $new_step = null : $new_step = $line['new_step'];
                
                $line['effective'] == '' ? $effective = null : $effective = $line['effective'];
                $line['placed'] == '' ? $placed = null : $placed = $line['placed'];
        
                $line['remark'] == '' ? $remark = null : $remark = $line['remark'];
                
                
                // $beneficiary = Beneficiary::find($ippis_no);

                if ($salary_structure == 'CONPASS') {

                    if((int)Carbon::instance($effective)->format('Y') >= 2019){
                        // VARIATION COMPUTATIONS
                        $old_salary_per_annum = Conpass::where('gl', $old_gl)
                        ->where('step', $old_step)
                        ->first()->salary_per_annum_with_shift;

                        $new_salary_per_annum = Conpass::where('gl', $new_gl)
                                            ->where('step', $new_step)
                                            ->first()->salary_per_annum_with_shift;
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
                        'dor' => $dor,
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
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'remark' => $remark,
                        ]);
                    }else{
                        // VARIATION COMPUTATIONS
                        $old_salary_per_annum = OldConpass::where('gl', $old_gl)
                        ->where('step', $old_step)
                        ->first()->salary_per_annum_with_shift;

                        $new_salary_per_annum = OldConpass::where('gl', $new_gl)
                                            ->where('step', $new_step)
                                            ->first()->salary_per_annum_with_shift;
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
                        'dor' => $dor,
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
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
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
                        'dor' => $dor,
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
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
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
                        'dor' => $dor,
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
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
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
                        'dor' => $dor,
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
                        'variation_amount' => $variation_amount,
                        'arrears' => $arrears,
                        'remark' => $remark,
                    ]);
                }

            });
            
            Alert::success('Variation imported successfully!', 'Success!')->autoclose(222500);
            return back();
        }
    }

    // GENERATE SINGLE REDEPLOYMENT SIGNAL
    public function generate_single_variation_letter($id, DNS2D $dNS2D){   

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

    // GENERATE BULK REDEPLOYMENT SIGNAL
    public function generate_bulk_variation_letter(Request $request, DNS2D $dNS2D){

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
                
                $image =$dNS2D->getBarcodePNG("
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
}