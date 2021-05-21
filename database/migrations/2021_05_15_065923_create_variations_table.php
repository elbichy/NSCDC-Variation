<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateVariationsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('variations', function (Blueprint $table) {
            $table->unsignedBigInteger('id')->autoIncrement();
            $table->string('svc_no')->nullable();
            $table->string('ippis_no')->nullable();
            $table->string('name')->nullable();
            $table->string('gender')->nullable();
            $table->dateTime('dob')->nullable();
            $table->string('present_rank')->nullable();
            $table->integer('present_gl')->nullable();
            $table->dateTime('dofa')->nullable();
            $table->string('salary_structure')->nullable();
            $table->string('paypoint')->nullable();
            $table->dateTime('dor')->nullable();

            $table->string('old_rank')->nullable();
            $table->integer('old_gl')->nullable();
            $table->integer('old_step')->nullable();
            $table->float('old_salary_per_annum', 20, 6)->nullable();
            $table->string('new_rank')->nullable();
            $table->integer('new_gl')->nullable();
            $table->integer('new_step')->nullable();
            $table->float('new_salary_per_annum', 20, 6)->nullable();
            $table->dateTime('effective')->nullable();
            $table->dateTime('placed')->nullable();

            $table->integer('months_owed')->nullable();
            $table->float('variation_amount', 20, 6)->nullable();
            $table->float('arrears', 20, 6)->nullable();
            $table->string('remark')->nullable();
            // $table->softDeletes();
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
        Schema::dropIfExists('variations');
    }
}
