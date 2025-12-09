<?php

namespace App\Exports;

use App\Exports\ReportDamageLinen\ReportDamageLinenRawSheetXlsxExport;
use App\Exports\ReportDamageLinen\ReportDamageLinenSheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ReportDamageXlsxExport implements WithMultipleSheets
{

    protected $HptCode;
    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }
    public function sheets(): array
    {
        $HptCode = $this->HptCode;
        return [
            new ReportDamageLinenSheetXlsxExport($HptCode),
            new ReportDamageLinenRawSheetXlsxExport($HptCode),
        ];
    }
}
