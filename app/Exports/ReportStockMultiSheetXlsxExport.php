<?php

namespace App\Exports;

use App\Exports\ReportStock\Stock1SheetXlsxExport;
use App\Exports\ReportStock\Stock2SheetXlsxExport;
use App\Exports\ReportStock\Stock3SheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ReportStockMultiSheetXlsxExport implements WithMultipleSheets
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
            new Stock1SheetXlsxExport($HptCode),
            new Stock2SheetXlsxExport($HptCode),
            new Stock3SheetXlsxExport($HptCode),
        ];
    }
}
