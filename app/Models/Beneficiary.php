<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Beneficiary extends Model
{
    use HasFactory;
    protected $guarded = [];

    protected $primaryKey = 'ippis_no';
    public $incrementing = false;
    protected $keyType = 'string';

    public function variations(){
        return $this->hasMany('App\Models\Variation', 'ippis_no', 'ippis_no');
    }
}
