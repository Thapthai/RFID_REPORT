<?php

namespace App\Exports\ReportUseLinen;

use Maatwebsite\Excel\Concerns\WithTitle;
use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithDrawings;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\WithEvents;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class ReportUseLinenSheetXlsxExport implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $DepCode;
    protected $DepName;
    protected $allDates;
    protected $dirtyAll;
    protected $cleanAll;
    protected $shelfcountAll;
    public function __construct(
        $DepCode,
        $DepName,
        $allDates,
        $dirtyAll,
        $cleanAll,
        $shelfcountAll
    ) {
        $this->DepCode = $DepCode;
        $this->DepName = $DepName;
        $this->allDates = $allDates;
        $this->dirtyAll = $dirtyAll;
        $this->cleanAll = $cleanAll;
        $this->shelfcountAll = $shelfcountAll;
    }

    public function title(): string
    {
        if (trim($this->DepName) === '') {
            return 'ไม่ได้ตั้งชื่อ';
        }

        return   $this->DepName;
    }

    public function view(): View
    {

        $DepCode = $this->DepCode;
        $DepName = $this->DepName;
        $allDates = $this->allDates;
        $dirtyAll = $this->dirtyAll;
        $shelfcountAll = $this->shelfcountAll;
        $cleanAll = $this->cleanAll;

        foreach ($allDates as $dateFormatted) {
            $key = $DepCode . '|' . $dateFormatted;

            $summaryData[] = [
                'DepCode' => $DepCode,
                'DepName' => $DepName,
                'date' => $dateFormatted,
                'dirtyLinen' => $dirtyAll[$key][0]->count ?? 0,
                'cleanLinen' => $cleanAll[$key][0]->count ?? 0,
                'Shelfcount' => $shelfcountAll[$key][0]->count ?? 0,
            ];
        }

        $reportDate =  date('m-Y');
        $thaiMonths = config('myconfig.thai_months');

        // แปลงเป็น timestamp
        $reportTimestamp = strtotime('01-' . $reportDate);
        $currentMonthNum = date('m', $reportTimestamp);

        $currentMonthName = $thaiMonths[$currentMonthNum];

        $currentYear = date('Y', $reportTimestamp);
        return view('exports.reportUseLinen_xlsx.reportUseLinenSheet', compact(
            'summaryData',
            'DepName',
            'currentMonthName',
            'currentYear'
        ));
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                $sheet->getPageSetup()->setScale(100); // บังคับให้ขนาดเหมือนกันทุกระบบ
                // $sheet->getPageSetup()->setFitToWidth(1);
                // $sheet->getPageSetup()->setFitToHeight(1);

                // ตั้งค่าขนาดกระดาษเป็น A4
                $sheet->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);
                $sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_PORTRAIT); // แนวตั้ง

                // // ตั้งค่าขอบกระดาษ (หน่วยเป็น points)
                $sheet->getPageMargins()->setTop(0.35);
                $sheet->getPageMargins()->setRight(0.25);
                $sheet->getPageMargins()->setLeft(0.35);
                $sheet->getPageMargins()->setBottom(0);

                $sheet->getStyle('B1')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 28,
                        'bold' => true
                    ],
                ]);

                $highestRow = $sheet->getHighestRow(); // หาบรรทัดสุดท้ายที่มีข้อมูล
                $sheet->getStyle("A2:Z{$highestRow}")->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 16,
                    ],
                ]);


                // ตั้งค่าความกว้างของคอลัมน์
                $columns = [
                    'A' => 25,
                    'B' => 25,
                    'C' => 25,
                    'D' => 25,

                ];
                foreach ($columns as $col => $width) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                // ปรับความสูงของแถว
                $rows = [
                    1 => 50,

                ];

                foreach ($rows as $row => $height) {
                    $sheet->getRowDimension($row)->setRowHeight($height);
                }

                $sheet->getStyle('B1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('B1')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

                $sheet->getStyle("A3:D{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle("A3:D{$highestRow}")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            },
        ];
    }

    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Company Logo');
        $drawing->setDescription('This is the company logo');
        $drawing->setPath(public_path('images/Nhealth_linen 4.0.png'));
        $drawing->setHeight(50);
        $drawing->setCoordinates('A1');
        $drawing->setOffsetX(30);
        $drawing->setOffsetY(10);

        return [$drawing];
    }
}
