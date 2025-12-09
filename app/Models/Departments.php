<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Departments extends Model
{
    use HasFactory;
    protected $table = 'department';
    protected $fillable = [
        'RowID',
        'DepCode',
        'HptCode',
        'DepName',
        'IsStatus',
        'IsDefault',
        'IsFactory',
        'DocDate',
        'Modify_Date',
        'Modify_Code',
        'IsActive',
        'GroupCode',
        'Ship_To',
    ];
}
