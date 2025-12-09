<?php

namespace App\Exports;

use App\Exports\ReportDamageLinenSelectType\ReportDamageLinenRawSelectTypeSheetXlsxExport;
use App\Exports\ReportDamageLinenSelectType\ReportDamageLinenSelectTypeSheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ReportDamageSelectTypeMultiSheetXlsxExport implements WithMultipleSheets
{
    protected $HptCode;
    protected $startDate;
    protected $endDate;
    protected $typeTH;
    public function __construct(
        $HptCode,
        $startDate,
        $endDate,
        $typeTH
    ) {
        $this->HptCode = $HptCode;
        $this->startDate = $startDate;
        $this->endDate = $endDate;
        $this->typeTH = $typeTH;
    }
    public function sheets(): array
    {
        $HptCode = $this->HptCode;
        $startDate = $this->startDate;
        $endDate = $this->endDate;
        $typeTH = $this->typeTH;

        $thaiMonths = config('myconfig.thai_months');
        $startDate = $this->startDate;
        $startDateDay = date('d',  strtotime($startDate));
        $startDateMonthNameTH = $thaiMonths[date('m',  strtotime($startDate))];
        $startDateYear =  date('Y', strtotime("+543 years", strtotime($startDate)));

        $endDate = $this->endDate;
        $endDateDay = date('d',  strtotime($endDate));
        $endDateMonthNameTH = $thaiMonths[date('m',  strtotime($endDate))];
        $endDateYear =  date('Y', strtotime("+543 years", strtotime($endDate)));

        $typeTopic = $typeTH . ' วันที่ ' . $startDateDay . ' ' . $startDateMonthNameTH . ' ' . $startDateYear .
            ($startDate != $endDate ? ' ถึง ' . $endDateDay . ' ' . $endDateMonthNameTH . ' ' . $endDateYear : '');

    
        return [
            new ReportDamageLinenSelectTypeSheetXlsxExport($HptCode, $startDate, $endDate, $typeTopic),
            new ReportDamageLinenRawSelectTypeSheetXlsxExport($HptCode, $startDate, $endDate, $typeTopic),
        ];
    }
}
