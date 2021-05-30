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
            [
                'service_number' => 66818,
                'fullname' => 'Suleiman Abdulrazaq',
                'email' =>'suleiman.bichi@gmail.com',
                'password' => Hash::make('@Suleimanu1'),
                'role' => 0,
            ],
            [
                'service_number' => 1212,
                'fullname' => 'Admin Officer',
                'email' =>'admin@gmail.com',
                'password' => Hash::make('@admin123'),
                'role' => 1,
            ],
            [
                'service_number' => 2323,
                'fullname' => 'Variation Officer',
                'email' =>'variation@gmail.com',
                'password' => Hash::make('@variation321'),
                'role' => 2,
            ],
            [
                'service_number' => 3333,
                'fullname' => 'IPPIS Officer',
                'email' =>'ippis@gmail.com',
                'password' => Hash::make('@ippis456'),
                'role' => 3,
            ],
        ]);
    }
}
