<?php

namespace App\Exports\ReportStock;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Maatwebsite\Excel\Concerns\WithTitle;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class Stock1SheetXlsxExport implements FromView, WithDrawings, WithEvents, WithTitle
{
    protected $HptCode;

    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }

    public function title(): string
    {
        return   "Stock 1";
    }

    public function view(): View
    {
        $HptCode = $this->HptCode;

        $reportDate =  date('m-Y');

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

        $data = DB::select(DB::raw("
        SELECT
          item.ItemName,
          itemstock_RFID.StatusDepartment,
          itemstock_RFID.StatusLaundry,
          itemstock_RFID.StatusLinenClean,
          factory.FacName,
          department.HptCode,
          department.DepName,
          COUNT(item.ItemName) AS Qty 
          FROM
          itemstock_RFID
          INNER JOIN item ON itemstock_RFID.ItemCode = item.ItemCode
          LEFT JOIN factory ON itemstock_RFID.IDfactory = factory.FacCode
          INNER JOIN department ON itemstock_RFID.DepCode = department.DepCode
          WHERE
          itemstock_RFID.HptCode = '$HptCode'
          AND DATEDIFF( DATE( NOW()), itemstock_RFID.LastDocDate ) < 30
          GROUP BY 
          -- itemstock_RFID.StatusDepartment , 
          -- itemstock_RFID.StatusLaundry , 
          -- itemstock_RFID.StatusLinenClean , 
          item.ItemName , 
          itemstock_RFID.DepCode 
          ORDER BY item.ItemName ASC
        "));

        // dd($data);
        return view('exports.reportStocks_xlsx.stock1', compact(
            'data',
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
                        'size' => 13,
                    ],
                ]);
                $sheet->getStyle('A2')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 28,
                        'bold' => true
                    ],
                ]);

                $highestRow = $sheet->getHighestRow(); // หาบรรทัดสุดท้ายที่มีข้อมูล
                $sheet->getStyle("A3:Z{$highestRow}")->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 16,
                    ],
                ]);


                // ตั้งค่าความกว้างของคอลัมน์
                $columns = [
                    'A' => 5,
                    'B' => 45,
                    'C' => 25,
                    'D' => 25,
                    'E' => 15,
                ];
                foreach ($columns as $col => $width) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                // ปรับความสูงของแถว
                $rows = [
                    2 => 50,

                ];

                foreach ($rows as $row => $height) {
                    $sheet->getRowDimension($row)->setRowHeight($height);
                }

                $sheet->getStyle('C1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $sheet->getStyle('C1')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

                $sheet->getStyle('A2:L3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A2:L3')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);


                $event->sheet->getDelegate()->getStyle('A3:E3')->applyFromArray([
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
        $drawing->setCoordinates('A2');
        $drawing->setOffsetX(60);
        $drawing->setOffsetY(10);

        return [$drawing];
    }
}
