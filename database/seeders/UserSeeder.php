<?php

namespace Database\Seeders;

use App\Models\User;
use Illuminate\Database\Seeder;
use Illuminate\Support\Facades\Hash;

class UserSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        User::insert([
            'service_number' => 66818,
            'fullname' => 'Suleiman Abdulrazaq',
            'email' =>'suleiman.bichi@gmail.com',
            'password' => Hash::make('@Suleimanu1'),
        ]);
    }
}
