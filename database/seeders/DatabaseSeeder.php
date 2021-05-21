<?php
namespace Database\Seeders;
use Illuminate\Database\Seeder;

class DatabaseSeeder extends Seeder
{
    /**
     * Seed the application's database.
     *
     * @return void
     */
    public function run()
    {
        $this->call([
            RankSeeder::class,
            ConpassSeeder::class,
            OldConpassSeeder::class,
            ConmessSeeder::class,
            ConhesspSeeder::class,
            ConhesshnSeeder::class,
        ]);
    }
}