<?php

namespace Database\Seeders;

use App\Models\Rank;
use Illuminate\Database\Seeder;

class RankSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        $ranks = array(
            array('id' => '1','cadre' => 'superintendent','gl' => '18','full_title' => 'Commandant General of Corps','short_title' => 'CG'),
            array('id' => '2','cadre' => 'superintendent','gl' => '17','full_title' => 'Deputy Commandant General of Corps','short_title' => 'DCG'),
            array('id' => '3','cadre' => 'superintendent','gl' => '16','full_title' => 'Assistant Commandant General of Corps','short_title' => 'ACG'),
            array('id' => '4','cadre' => 'superintendent','gl' => '15','full_title' => 'Commandant of Corps','short_title' => 'CC'),
            array('id' => '5','cadre' => 'superintendent','gl' => '14','full_title' => 'Deputy Commandant of Corps','short_title' => 'DCC'),
            array('id' => '6','cadre' => 'superintendent','gl' => '13','full_title' => 'Assistant Commandant of Corps','short_title' => 'ACC'),
            array('id' => '7','cadre' => 'superintendent','gl' => '12','full_title' => 'Chief Superintendent of Corps','short_title' => 'CSC'),
            array('id' => '8','cadre' => 'superintendent','gl' => '11','full_title' => 'Superintendent of Corps','short_title' => 'SC'),
            array('id' => '9','cadre' => 'superintendent','gl' => '10','full_title' => 'Deputy Superintendent of Corps','short_title' => 'DSC'),
            array('id' => '10','cadre' => 'superintendent','gl' => '9','full_title' => 'Assistant Superintendent of Corps I','short_title' => 'ASC I'),
            array('id' => '11','cadre' => 'superintendent','gl' => '8','full_title' => 'Assistant Superintendent of Corps II','short_title' => 'ASC II'),
            
            array('id' => '12','cadre' => 'inspectorate','gl' => '7','full_title' => 'Inspector of Corps/Chief Corps Assistant','short_title' => 'IC/CCA'),
            array('id' => '13','cadre' => 'inspectorate','gl' => '6','full_title' => 'Assistant Inspector of Corps/Senior Corps Assistant','short_title' => 'AIC/SCA'),
            array('id' => '14','cadre' => 'assistant','gl' => '5','full_title' => 'Corps Assistant I','short_title' => 'CA I'),
            array('id' => '15','cadre' => 'assistant','gl' => '4','full_title' => 'Corps Assistant II','short_title' => 'CA II'),
            array('id' => '16','cadre' => 'assistant','gl' => '3','full_title' => 'Corps Assistant III','short_title' => 'CA III')
        );
        
        foreach ($ranks as $key => $value) {
            Rank::insert($value);
        }
    }
}
