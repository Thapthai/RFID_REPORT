<?php

namespace App\Exports;

use App\Exports\ReportDamageLinenDaily\ReportDamageLinenDailySheetXlsxExport;
use App\Exports\ReportDamageLinenDaily\ReportDamageLinenRawDailySheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ReportDamageDailyMultiSheetXlsxExport implements WithMultipleSheets
{
    protected $HptCode;
    protected $DocDate;
    public function __construct($HptCode, $DocDate)
    {
        $this->HptCode = $HptCode;
        $this->DocDate = $DocDate;
    }
    public function sheets(): array
    {
        $HptCode = $this->HptCode;
        $DocDate = $this->DocDate;
        return [
            new ReportDamageLinenDailySheetXlsxExport($HptCode, $DocDate),
            new ReportDamageLinenRawDailySheetXlsxExport($HptCode, $DocDate),
        ];
    }
}
