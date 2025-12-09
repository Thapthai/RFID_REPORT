<?php

namespace App\Exports\ReportStock;

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
use PhpOffice\PhpSpreadsheet\Style\Fill;


class Stock2SheetXlsxExport implements FromArray, WithHeadings, WithEvents, WithDrawings, WithCustomStartCell, WithTitle
{
    protected $HptCode;
    protected $hiddenRows = [];

    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }

    public function title(): string
    {
        return   "Stock 2";
    }
    public function startCell(): string
    {
        return 'A3';
    }

    public function array(): array
    {
        $rows = [];
        $rowPointer = 4;

        $mainItems = DB::select("
            SELECT item.ItemName, item.ItemCode, COUNT(item.ItemName) AS Qty
            FROM itemstock_RFID
            INNER JOIN item ON itemstock_RFID.ItemCode = item.ItemCode
            LEFT JOIN factory ON itemstock_RFID.IDfactory = factory.FacCode
            INNER JOIN department ON itemstock_RFID.DepCode = department.DepCode
            WHERE itemstock_RFID.HptCode = ? 
              AND DATEDIFF(DATE(NOW()), itemstock_RFID.LastDocDate) < 30
            GROUP BY item.ItemCode 
            ORDER BY item.ItemName
        ", [$this->HptCode]);

        $count = 1;

        foreach ($mainItems as $item) {
            // Main row
            $rows[] = [$count++, $item->ItemName, $item->Qty];
            $rowPointer++;

            $subItems = DB::select("
                SELECT itemstock_RFID.StatusDepartment, itemstock_RFID.StatusLaundry,
                       itemstock_RFID.StatusLinenClean, factory.FacName, department.DepName,
                       COUNT(item.ItemName) AS Qty
                FROM itemstock_RFID
                INNER JOIN item ON itemstock_RFID.ItemCode = item.ItemCode
                LEFT JOIN factory ON itemstock_RFID.IDfactory = factory.FacCode
                INNER JOIN department ON itemstock_RFID.DepCode = department.DepCode
                WHERE itemstock_RFID.HptCode = ? AND item.ItemCode = ? 
                  AND DATEDIFF(DATE(NOW()), itemstock_RFID.LastDocDate) < 30
                GROUP BY itemstock_RFID.DepCode, item.ItemName
                ORDER BY item.ItemName 
            ", [$this->HptCode, $item->ItemCode]);

            foreach ($subItems as $sub) {
                $location = $sub->StatusDepartment ? $sub->DepName : ($sub->StatusLaundry ? $sub->FacName : ($sub->StatusLinenClean ? $sub->DepName : ''));
                $rows[] = ['', $location, $sub->Qty];
                $this->hiddenRows[] = $rowPointer++;
            }
            $rows[] = ['', '', ''];
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
                $sheet->setCellValue('A2', 'รายงานสต๊อกคงคลัง');
                $sheet->setCellValue('C1', 'วันที่พิมพ์รายงาน ' . date('d') . ' ' . $currentMonthName . ' พ.ศ. ' . ($currentYear + 543));

                // ฟอนต์ของหัวข้อบนสุด
                $sheet->getStyle('C1')->applyFromArray([
                    'font' => ['name' => 'Angsana New', 'size' => 13],
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

                // Style ทั่วไปในพื้นที่ข้อมูล
                $highestRow = $sheet->getHighestRow();
                $sheet->getStyle("A3:Z{$highestRow}")->applyFromArray([
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
                        'A' => 5,
                        'B' => 45,
                        'C' => 25,

                    ] as $col => $width
                ) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                // ความสูงของแถว
                $sheet->getRowDimension(1)->setRowHeight(20);
                $sheet->getRowDimension(2)->setRowHeight(40);

                // เส้นขอบบาง
                $sheet->getStyle("A3:C{$highestRow}")->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                ]);

                $event->sheet->getDelegate()->getStyle('A3:C3')->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => 'A9DDFC',
                        ],
                    ],
                ]);

                // ซ่อน subData (กรณีใช้ outline)
                if (property_exists($this, 'hiddenRows') && is_array($this->hiddenRows)) {
                    foreach ($this->hiddenRows as $row) {
                        $sheet->getRowDimension($row)->setVisible(false);
                        $sheet->getRowDimension($row)->setCollapsed(true);
                        $sheet->getRowDimension($row)->setOutlineLevel(1);
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
