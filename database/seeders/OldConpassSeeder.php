<?php

namespace Database\Seeders;

use App\Models\OldConpass;
use Rap2hpoutre\FastExcel\FastExcel;
use Illuminate\Database\Seeder;

class OldConpassSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        $collection = (new FastExcel)->import(storage_path('old_conpass_salary_chart.csv'));
        foreach ($collection as $row) {
            // dd((float)$row['net_pay'];
            $step = OldConpass::create([
                'rank_id' => $row['rank_id'] == '' ? null : (float)$row['rank_id'],
                'gl' => $row['gl'] == '' ? null : (float)$row['gl'],
                'step' => $row['step'] == '' ? null : (float)$row['step'],
                'salary_per_annum_with_shift' => $row['salary_per_annum_with_shift'] == '' ? null : (float)$row['salary_per_annum_with_shift'],
                'salary_per_annum' => $row['salary_per_annum'] == '' ? null : (float)$row['salary_per_annum'],
                // 'rent_per_annum' => $row['rent_per_annum'] == '' ? null : (float)$row['rent_per_annum'],
                'monthly_salary' => $row['monthly_salary'] == '' ? null : (float)$row['monthly_salary'],
                'monthly_rent' => $row['monthly_rent'] == '' ? null : (float)$row['monthly_rent'],
                'shift_duty_allowance' => $row['shift_duty_allowance'] == '' ? null : (float)$row['shift_duty_allowance'],
                'domestic_servant' => $row['domestic_servant'] == '' ? null : (float)$row['domestic_servant'],
                'gross_emolument' => $row['gross_emolument'] == '' ? null : (float)$row['gross_emolument'],
                'tax' => $row['tax'] == '' ? null : (float)$row['tax'],
                'pension' => $row['pension'] == '' ? null : (float)$row['pension'],
                'nhf' => $row['nhf'] == '' ? null : (float)$row['nhf'],
                'total_deduction' => $row['total_deduction'] == '' ? null : (float)$row['total_deduction'],
                'net_pay' => $row['net_pay'] == '' ? null : (float)$row['net_pay'],
            ]);
        }
    }
}
