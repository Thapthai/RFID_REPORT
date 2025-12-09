<?php

namespace App\Exports\ReportDamageLinenDaily;

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

class ReportDamageLinenDailySheetXlsxExport implements FromArray, WithHeadings, WithEvents, WithDrawings, WithCustomStartCell, WithTitle
{
    protected $HptCode;
    protected $DocDate;
    protected $docDateRows = [];
    protected $itemRows = [];


    public function __construct($HptCode, $DocDate)
    {
        $this->HptCode = $HptCode;
        $this->DocDate = $DocDate;
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

        // dd($count);
        $DocDate = $this->DocDate;

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
                WHERE   
                    damagenh.DepCode = 'BPHCENTER' 
                    AND damagenh_detail_round.DepCode != ''
                    AND DATE(damagenh.DocDate) = ?

                GROUP BY department.DepName, department.DepCode
            ", [$DocDate]);

        foreach ($mainItems as $item) {
            // Main row
            $rows[] = [$count++, $item->DepName, '', ''];
            $rowPointer++;

            // ดึง DocDate ที่เกี่ยวข้องกับ Dep นี้
            $docDates = DB::select("
                SELECT DISTINCT DATE(damagenh.DocDate) AS DocDate
                FROM damagenh
                INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
                INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
                INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
                WHERE damagenh.DepCode = 'BPHCENTER'
                    AND department.DepCode = ?
                    AND DATE(damagenh.DocDate) = ?
                ORDER BY damagenh.DocDate
            ", [$item->DepCode, $DocDate]);

            foreach ($docDates as $doc) {
                $rows[] = ['', 'วันที่ ' . date('d-m-Y', strtotime("+543 years", strtotime($doc->DocDate))), '', ''];

                $this->docDateRows[] = $rowPointer++; // สำหรับ outline level 1

                // ดึงรายการ item ตาม DocDate นี้
                $subItems = DB::select("
                    SELECT
                        item.ItemName,
                        COUNT(itemstock_RFID.RfidCode) AS qty
                    FROM
                        damagenh
                        INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
                        INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
                        INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
                        INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
                        INNER JOIN itemstock_RFID ON damagenh_detail_round.RFID = itemstock_RFID.RfidCode 
                    WHERE
                        damagenh.DepCode = 'BPHCENTER' 
                        AND department.DepCode = ?
                        AND DATE(damagenh.DocDate) = ?
                    GROUP BY item.ItemName
                    ORDER BY item.ItemName
                ", [$item->DepCode, $doc->DocDate]);

                foreach ($subItems as $sub) {
                    $rows[] = ['', $sub->ItemName, $sub->qty];
                    $this->itemRows[] = $rowPointer++; // สำหรับ outline level 2
                }
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
                    'font' => ['name' => 'Angsana New', 'size' => 28, 'bold' => true],
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

                // ซ่อน docDateRows (level 1)
                if (property_exists($this, 'docDateRows') && is_array($this->docDateRows)) {
                    foreach ($this->docDateRows as $row) {
                        $sheet->getRowDimension($row)->setVisible(false);
                        $sheet->getRowDimension($row)->setCollapsed(true);
                        $sheet->getRowDimension($row)->setOutlineLevel(1);

                        $event->sheet->getDelegate()->getStyle("A{$row}:C{$row}")->applyFromArray([
                            'fill' => [
                                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                                'startColor' => [
                                    'rgb' => 'EEEEEE', // สี #eee
                                ],
                            ],
                        ]);
                    }
                }

                // ซ่อน itemRows (level 2)
                if (property_exists($this, 'itemRows') && is_array($this->itemRows)) {
                    foreach ($this->itemRows as $row) {
                        $sheet->getRowDimension($row)->setVisible(false);
                        $sheet->getRowDimension($row)->setCollapsed(true);
                        $sheet->getRowDimension($row)->setOutlineLevel(2);
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
