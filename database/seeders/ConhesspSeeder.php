<?php

namespace Database\Seeders;

use App\Models\Conhessp;
use Illuminate\Database\Seeder;
use Rap2hpoutre\FastExcel\FastExcel;

class ConhesspSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        $collection = (new FastExcel)->import(storage_path('conhessp_salary_chart.csv'));
        foreach ($collection as $row) {
            // dd((float)$row['net_pay'];
            $step = Conhessp::create([
                'rank_id' => $row['rank_id'] == '' ? null : (float)$row['rank_id'],
                'conhessp_gl' => $row['conhessp_gl'] == '' ? null : (float)$row['conhessp_gl'],
                'conhessp_step' => $row['conhessp_step'] == '' ? null : (float)$row['conhessp_step'],
                'conpass_gl' => $row['conpass_gl'] == '' ? null : (float)$row['conpass_gl'],
                'conpass_step' => $row['conpass_step'] == '' ? null : (float)$row['conpass_step'],

                'consolidated_salary_per_annum' => $row['consolidated_salary_per_annum'] == '' ? null : (float)$row['consolidated_salary_per_annum'],
                'monthly_consolidated_salary' => $row['monthly_consolidated_salary'] == '' ? null : (float)$row['monthly_consolidated_salary'],
                'non_clinic_allowance' => $row['non_clinic_allowance'] == '' ? null : (float)$row['non_clinic_allowance'],
                'call_duty' => $row['call_duty'] == '' ? null : (float)$row['call_duty'],
                'hazard_allowance' => $row['hazard_allowance'] == '' ? null : (float)$row['hazard_allowance'],
                'gross_emolument' => $row['gross_emolument'] == '' ? null : (float)$row['gross_emolument'],
                'tax' => $row['tax'] == '' ? null : (float)$row['tax'],
                'nhf' => $row['nhf'] == '' ? null : (float)$row['nhf'],
                'pension' => $row['pension'] == '' ? null : (float)$row['pension'],
                'total_deduction' => $row['total_deduction'] == '' ? null : (float)$row['total_deduction'],
                'net_pay' => $row['net_pay'] == '' ? null : (float)$row['net_pay'],
            ]);
        }
    }
}
