<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateConhesshnsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('conhesshns', function (Blueprint $table) {
            $table->unsignedBigInteger('id')->autoIncrement();
            $table->bigInteger('rank_id')->nullable();
            $table->bigInteger('conhesshn_gl')->nullable();
            $table->bigInteger('conhesshn_step')->nullable();
            $table->bigInteger('conpass_gl')->nullable();
            $table->bigInteger('conpass_step')->nullable();
            $table->float('consolidated_salary_per_annum', 20, 6)->nullable();
            $table->float('monthly_consolidated_salary', 20, 6)->nullable();
            $table->float('non_clinic_allowance', 20, 6)->nullable();
            $table->float('call_duty', 20, 6)->nullable();
            $table->float('hazard_allowance', 20, 6)->nullable();
            $table->float('gross_emolument')->nullable();
            $table->float('tax', 20, 6);
            $table->float('nhf', 20, 6)->nullable();
            $table->float('pension', 20, 6)->nullable();
            $table->float('total_deduction', 20, 6)->nullable();
            $table->float('net_pay', 20, 6)->nullable();
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('conhesshns');
    }
}
