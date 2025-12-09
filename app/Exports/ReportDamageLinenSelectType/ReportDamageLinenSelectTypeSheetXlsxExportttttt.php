<?php

namespace App\Exports\ReportDamageLinenSelectType;

use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithCustomStartCell;

class ReportDamageLinenSelectTypeSheetXlsxExportttttt implements FromArray, WithHeadings, WithEvents, WithDrawings, WithCustomStartCell, WithTitle
{
    protected $HptCode;
    protected $startDate;
    protected $endDate;
    protected $typeTopic;
    protected $docDateRows = [];
    protected $itemRows = [];
    public function __construct($HptCode, $startDate, $endDate, $typeTopic)
    {
        $this->HptCode = $HptCode;
        $this->startDate = $startDate;
        $this->endDate = $endDate;
        $this->typeTopic = $typeTopic;
    }

    public function title(): string
    {
        return   "รายการผ้าชำรุด";
    }
    public function startCell(): string
    {
        return 'A4';
    }

    public function array(): array
    {
        $rows = [];
        $rowPointer = 5;
        $count = 1;

        $startDate = $this->startDate;
        $endDate = $this->endDate;

    $docDateRows = [];
    $itemRows = [];
        $mainItems = DB::select("
            SELECT
                item.ItemName,
                item.ItemCode,
                COUNT(item.ItemName) AS qty
            FROM
                stock_outstandingdelivery
            INNER JOIN item ON stock_outstandingdelivery.ItemCode = item.ItemCode
            INNER JOIN department ON stock_outstandingdelivery.DepCode = department.DepCode
            WHERE
                stock_outstandingdelivery.IsStatus = 0
                AND stock_outstandingdelivery.DepCode = 'ff1'
            GROUP BY
                item.ItemName, stock_outstandingdelivery.DocDate
            ORDER BY
                item.ItemName ASC;

            ");

        foreach ($mainItems as $item) {
            // Main row
            $rows[] = [$count++, $item->ItemName, $item->qty, ''];
            $rowPointer++;
            $subitems = DB::select("
            SELECT
                item.ItemName,
                stock_outstandingdelivery.RefDocNo,
                stock_outstandingdelivery.DocDate,
                stock_outstandingdelivery.RFID,
                stock_outstandingdelivery.QrCode
                COUNT(item.ItemName) AS qty

            FROM
                stock_outstandingdelivery
            INNER JOIN item ON stock_outstandingdelivery.ItemCode = item.ItemCode
            INNER JOIN department ON stock_outstandingdelivery.DepCode = department.DepCode
            WHERE stock_outstandingdelivery.IsStatus = 0  
            AND stock_outstandingdelivery.DepCode = 'ff1' 
            AND item.ItemCode = ?
            ORDER BY item.ItemName ASC

            ", [$item->ItemCode]);

            foreach ($subitems as $subitem) {
                $rows[] = ['', $subitem->ItemName,$subitem->DocDate,$subitem->qty];

                $this->docDateRows[] = $rowPointer++; // สำหรับ outline level 1
            }
            $rows[] = ['', '', '', ''];
            $rowPointer++;
        }

        return $rows;
    }


    public function headings(): array
    {
        return ['ลำดับ', 'รายการ', 'จำนวน'];
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {

                $reportDate =  date('m-Y');

                $thaiMonths = config('myconfig.thai_months');

                // แปลงเป็น timestamp
                $reportTimestamp = strtotime('01-' . $reportDate);
                $currentMonthNum = date('m', $reportTimestamp);
                $currentMonthName = $thaiMonths[$currentMonthNum];
                $currentYear = date('Y', $reportTimestamp);

                $sheet = $event->sheet->getDelegate();

                // ขนาดกระดาษ และการพิมพ์
                $sheet->getPageSetup()
                    ->setScale(100)
                    ->setPaperSize(PageSetup::PAPERSIZE_A4)
                    ->setOrientation(PageSetup::ORIENTATION_PORTRAIT);

                // ขอบกระดาษ
                $sheet->getPageMargins()
                    ->setTop(0.35)
                    ->setRight(0.25)
                    ->setLeft(0.35)
                    ->setBottom(0.3);

                $sheet->mergeCells('A2:C2');
                $sheet->setCellValue('A2', 'รายงานผ้าชำรุด');
                $sheet->mergeCells('A3:C3');
                $sheet->setCellValue('C1', 'วันที่พิมพ์รายงาน ' . date('d') . ' ' . $currentMonthName . ' พ.ศ. ' . ($currentYear + 543));
                $sheet->setCellValue('A3', $this->typeTopic);

                // ฟอนต์ของหัวข้อบนสุด
                $sheet->getStyle('C1')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 13
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_RIGHT,
                        'vertical' => Alignment::VERTICAL_TOP,
                    ],
                ]);

                // หัวเรื่องหลัก
                $sheet->getStyle('A2')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 28,
                        'bold' => true
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_CENTER,
                        'vertical' => Alignment::VERTICAL_CENTER,
                    ],
                ]);

                $sheet->getStyle('A3')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 20
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_CENTER,
                        'vertical' => Alignment::VERTICAL_CENTER,
                    ],
                ]);

                $event->sheet->getDelegate()->getStyle('A4:C4')->applyFromArray([
                    'fill' => [
                        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => '8fdaff',
                        ],
                    ],
                ]);

                // Style ทั่วไปในพื้นที่ข้อมูล
                $highestRow = $sheet->getHighestRow();
                $sheet->getStyle("A4:Z{$highestRow}")->applyFromArray([
                    'font' => ['name' => 'Angsana New', 'size' => 16],
                    'alignment' => [
                        'vertical' => Alignment::VERTICAL_CENTER,
                    ],
                ]);

                // ตรงกลางหัวตาราง
                $sheet->getStyle('A2:L3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle("C4:C{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle("A4:A{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

                // ความกว้างของคอลัมน์
                foreach (
                    [
                        'A' => 10,
                        'B' => 60,
                        'C' => 25,

                    ] as $col => $width
                ) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                // ความสูงของแถว
                $sheet->getRowDimension(1)->setRowHeight(20);
                $sheet->getRowDimension(2)->setRowHeight(40);
                $sheet->getRowDimension(3)->setRowHeight(30);

                // เส้นขอบบาง
                $sheet->getStyle("A3:C{$highestRow}")->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                ]);

                // dd($this->itemRows);


                // ซ่อน docDateRows (level 1)
                if (property_exists($this, 'docDateRows') && is_array($this->docDateRows)) {
                    foreach ($this->docDateRows as $index => $row) {

                        $sheet->getRowDimension($row)->setOutlineLevel(1)
                            ->setVisible(false)
                            ->setCollapsed(true);

                        $event->sheet->getDelegate()->getStyle("A{$row}:C{$row}")->applyFromArray([
                            'fill' => [
                                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                                'startColor' => [
                                    'rgb' => 'fff1bf',
                                ],
                            ],
                        ]);
                    }
                }
                // ซ่อน itemRows (level 2)
                if (property_exists($this, 'itemRows') && is_array($this->itemRows)) {
                    foreach ($this->itemRows as $row) {
                        $sheet->getRowDimension($row)->setOutlineLevel(2)
                            ->setVisible(false)
                            ->setCollapsed(true);
                    }
                }

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
        $drawing->setOffsetX(10);
        $drawing->setOffsetY(10);

        return [$drawing];
    }
}
