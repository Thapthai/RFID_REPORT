<?php

namespace App\Exports;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Maatwebsite\Excel\Concerns\WithTitle;

class ReportStockBalanceXlsxExport implements FromView, WithDrawings, WithEvents, WithTitle
{
    protected $data;
    protected $reportDate;

    public function __construct($data, $reportDate)
    {
        $this->data = $data;
        $this->reportDate = $reportDate;
    }

    public function title(): string
    {
        return   "Stock Balance";
    }

    public function view(): View
    {
        $data = $this->data;
        $reportDate =  $this->reportDate;

        $thaiMonths = config('myconfig.thai_months');

        // แปลงเป็น timestamp
        $reportTimestamp = strtotime('01-' . $reportDate);
        $previousTimestamp = strtotime('-1 month', $reportTimestamp);
        $currentMonthNum = date('m', $reportTimestamp);
        $previousMonthNum = date('m', $previousTimestamp);

        $currentMonthName = $thaiMonths[$currentMonthNum];
        $previousMonthName = $thaiMonths[$previousMonthNum];

        $currentYear = date('Y', $reportTimestamp);
        $previousYear = date('Y', $previousTimestamp);

        return view('exports.report_stock_balance_xlsx', compact(
            'data',
            'reportDate',
            'currentMonthName',
            'previousMonthName',
            'currentYear',
            'previousYear'
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

                $sheet->getStyle('C1')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 28,
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
                    'A' => 5,
                    'B' => 45,
                    'C' => 20,
                    'D' => 20,
                    'E' => 20,
                    'F' => 20,
                    'G' => 15,
                    'H' => 15,
                    'I' => 15,
                    'J' => 15,
                    'K' => 15,
                    'L' => 15,

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

                $sheet->getStyle('C1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('C1')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

                $sheet->getStyle('A2:L4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A2:L4')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

                $event->sheet->getDelegate()->getStyle('A3:L4')->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => 'A9DDFC',
                        ],
                    ],
                ]);
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
        $drawing->setOffsetX(60);
        $drawing->setOffsetY(10);

        return [$drawing];
    }
}
