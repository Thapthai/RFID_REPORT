<?php

namespace App\Exports;

use App\Exports\ReportUseLinen\ReportUseLinenSheetXlsxExport;
use App\Exports\ReportUseLinen\ReportUseLinenSummarySheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ReportUseLinenMultiSheetXlsxExport implements WithMultipleSheets
{
    protected $departments;
    protected $depCode;
    protected $allDates;
    protected $dirtyAll;
    protected $cleanAll;
    protected $shelfcountAll;

    public function __construct(
        $departments,
        $depCode,
        $allDates,
        $dirtyAll,
        $cleanAll,
        $shelfcountAll
    ) {
        $this->departments = $departments;
        $this->depCode = $depCode;
        $this->allDates = $allDates;
        $this->dirtyAll = $dirtyAll;
        $this->cleanAll = $cleanAll;
        $this->shelfcountAll = $shelfcountAll;
    }
    public function sheets(): array
    {
        $departments = $this->departments;
        $allDates = $this->allDates;
        $dirtyAll = $this->dirtyAll;
        $cleanAll = $this->cleanAll;
        $shelfcountAll = $this->shelfcountAll;

        $sheets = [];

        if ((string)$this->depCode  === '0') {
            $sheets[] = new ReportUseLinenSummarySheetXlsxExport();
        }
 

        foreach ($departments as $DepCode => $DepName) {
            $sheets[] = new ReportUseLinenSheetXlsxExport(
                $DepCode,
                $DepName,
                $allDates,
                $dirtyAll,
                $cleanAll,
                $shelfcountAll,
            );
        }

        return $sheets;
    }
}
