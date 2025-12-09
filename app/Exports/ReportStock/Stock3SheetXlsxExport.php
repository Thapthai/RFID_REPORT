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
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Concerns\WithTitle;
use PhpOffice\PhpSpreadsheet\Style\Fill;


class Stock3SheetXlsxExport implements FromView, WithDrawings, WithEvents, WithTitle
{
    protected $HptCode;

    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }

    public function title(): string
    {
        return   "Stock 3";
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
                    SELECT item.ItemName, item.ItemCode,department.DepName  AS xName,
                    (SELECT Count(it.id) FROM itemstock_RFID it WHERE it.ItemCode = itemstock_RFID.ItemCode AND it.HptCode = '$HptCode') AS cntAll,
                            SUM(itemstock_RFID.StatusLaundry = 1  AND itemstock_RFID.LastDocNo != 'CNPOS2000-00001') AS cntDirty,
                            SUM(itemstock_RFID.StatusLinenClean = 1  AND itemstock_RFID.LastDocNo != 'CNPOS2000-00001') AS cntClean,
                            SUM(itemstock_RFID.StatusDepartment = 1  AND itemstock_RFID.LastDocNo != 'CNPOS2000-00001') AS cntSticker,																			
                    (SELECT Count(it.isCancel) FROM itemstock_RFID it WHERE it.isCancel = 1 AND it.ItemCode = itemstock_RFID.ItemCode AND it.HptCode = '$HptCode') AS cntDam,

                    CASE WHEN (SELECT Count( it.RfidCode ) FROM itemstock_RFID it WHERE it.StatusLinenClean = 1 
                        AND DATEDIFF( DATE( NOW()), it.LastDocDate ) > site.alertmotion AND it.ItemCode = item.ItemCode) 
                    THEN 'text-danger' ELSE '' END AS Clean,

                    CASE WHEN (SELECT Count( it.RfidCode ) FROM itemstock_RFID it WHERE it.StatusLaundry = 1 
                        AND DATEDIFF( DATE( NOW()), it.LastDocDate ) > site.alertmotion AND it.ItemCode = item.ItemCode) 
                    THEN 'text-danger' ELSE '' END AS drity,

                    CASE WHEN (SELECT Count( it.RfidCode ) FROM itemstock_RFID it WHERE it.StatusDepartment = 1 
                        AND DATEDIFF( DATE( NOW()), it.LastDocDate ) > site.alertmotion AND it.ItemCode = item.ItemCode ) 
                    THEN 'text-danger' ELSE '' END AS Dep

                    FROM itemstock_RFID
                        INNER JOIN  item ON itemstock_RFID.ItemCode = item.ItemCode
                        INNER JOIN  department ON itemstock_RFID.DepCode = department.DepCode
                        INNER JOIN  site ON itemstock_RFID.HptCode = site.HptCode
                        
                    WHERE itemstock_RFID.isCancel = 0 AND itemstock_RFID.HptCode = '$HptCode' 
                    GROUP BY itemstock_RFID.ItemCode 
                    ORDER BY item.ItemName ASC
                "));

        // dd($data);

        return view('exports.reportStocks_xlsx.stock3', compact(
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
                    'C' => 20,
                    'D' => 20,
                    'E' => 20,
                    'F' => 20,
                    'G' => 20
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

                $event->sheet->getDelegate()->getStyle('A3:G3')->applyFromArray([
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
