<?php

namespace App\Exports\ReportDamageLinen;

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

class ReportDamageLinenSheetXlsxExport implements FromArray, WithHeadings, WithEvents, WithDrawings, WithCustomStartCell, WithTitle
{
    protected $HptCode;
    protected $hiddenRows = [];

    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }

    public function title(): string
    {
        return   "รายการผ้าชำรุด";
    }
    public function startCell(): string
    {
        return 'A3';
    }

    public function array(): array
    {
        $rows = [];
        $rowPointer = 4;
        $count = 1;

        // Main items query: group by DepName และ DepCode พร้อมแสดงจำนวนรวม
        $mainItems = DB::select("
                SELECT
                    department.DepName,
                    department.DepCode,
                    COUNT(itemstock_RFID.RfidCode) AS qty
                FROM damagenh
                INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
                INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
                LEFT JOIN department ON damagenh_detail_round.DepCode = department.DepCode
                INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
                INNER JOIN itemstock_RFID ON damagenh_detail_round.RFID = itemstock_RFID.RfidCode 
                WHERE damagenh.IsStatus = 1 
                    -- AND damagenh.DepCode = 'BPHCENTER' 
                AND damagenh_detail_round.DepCode != '' 
                GROUP BY department.DepName, department.DepCode
            ");

        foreach ($mainItems as $item) {
            // Main row
            $rows[] = [$count++, $item->DepName, $item->qty];
            $rowPointer++;

            // Sub-items query
            $subItems = DB::select("
                    SELECT
                    item.ItemName,
                    itemstock_RFID.RfidCode,
                    itemstock_RFID.QrCode,
                    COUNT(itemstock_RFID.RfidCode) AS qty,
                    damagenh.DocDate 
                FROM
                    damagenh
                    INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
                    INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
                    INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
                    INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
                    INNER JOIN itemstock_RFID ON damagenh_detail_round.RFID = itemstock_RFID.RfidCode 
                WHERE
                    department.DepCode = ?
                    -- AND damagenh.DepCode = 'BPHCENTER' 
                    AND  damagenh.IsStatus = 1 
                  GROUP BY item.ItemName ASC
                ", [$item->DepCode]);

            foreach ($subItems as $sub) {
                $itemName = $sub->ItemName;
                $qty = $sub->qty;
                $rows[] = ['', $itemName, $qty]; // ใช้ qty ที่นับจาก DB
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
                $sheet->setCellValue('A2', 'รายงานผ้าชำรุด');
                $sheet->setCellValue('C1', 'วันที่พิมพ์รายงาน ' . date('d') . ' ' . $currentMonthName . ' พ.ศ. ' . ($currentYear + 543));

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

                $event->sheet->getDelegate()->getStyle('A3:C3')->applyFromArray([
                    'fill' => [
                        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => '8fdaff',
                        ],
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

                // เส้นขอบบาง
                $sheet->getStyle("A3:C{$highestRow}")->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => 'FF000000'],
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
