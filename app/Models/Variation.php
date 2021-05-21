<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Variation extends Model
{
    use HasFactory;
    protected $guarded = [];

    public function beneficiary(){
        return $this->belongsTo('App\Models\Beneficiary', 'ippis_no', 'ippis_no');
    }
}
