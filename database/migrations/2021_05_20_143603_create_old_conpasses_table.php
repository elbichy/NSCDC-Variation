<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateOldConpassesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('old_conpasses', function (Blueprint $table) {
            $table->unsignedBigInteger('id')->autoIncrement();
            $table->bigInteger('rank_id')->nullable();
            $table->bigInteger('gl')->nullable();
            $table->bigInteger('step')->nullable();
            $table->float('salary_per_annum_with_shift', 20, 6)->nullable();
            $table->float('salary_per_annum', 20, 6)->nullable();
            // $table->float('rent_per_annum', 20, 6)->nullable();
            $table->float('monthly_salary', 20, 6)->nullable();
            $table->float('monthly_rent', 20, 6)->nullable();
            $table->float('shift_duty_allowance', 20, 6)->nullable();
            $table->float('domestic_servant', 20, 6)->nullable();
            $table->float('gross_emolument', 20, 6)->nullable();
            $table->float('tax', 20, 6);
            $table->float('pension', 20, 6)->nullable();
            $table->float('nhf', 20, 6)->nullable();
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
        Schema::dropIfExists('old_conpasses');
    }
}
