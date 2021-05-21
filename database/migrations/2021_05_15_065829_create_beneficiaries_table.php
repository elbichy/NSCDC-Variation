<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateBeneficiariesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('beneficiaries', function (Blueprint $table) {
            $table->string('svc_no')->nullable();
            $table->string('ippis_no')->nullable()->primary();
            $table->string('name')->nullable();
            $table->string('gender')->nullable();
            $table->dateTime('dob')->nullable();
            $table->string('rank')->nullable();
            $table->integer('gl')->nullable();
            $table->dateTime('dofa')->nullable();
            $table->string('salary_structure')->nullable();
            $table->dateTime('dor')->nullable();
            $table->string('paypoint')->nullable();
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
        Schema::dropIfExists('beneficiaries');
    }
}
