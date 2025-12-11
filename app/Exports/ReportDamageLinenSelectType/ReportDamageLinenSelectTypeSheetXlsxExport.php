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

class ReportDamageLinenSelectTypeSheetXlsxExport implements FromArray, WithHeadings, WithEvents, WithDrawings, WithCustomStartCell, WithTitle
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


        // Main items query: group by DepName และ DepCode พร้อมแสดงจำนวนรวม
        $mainItems = DB::select("
        SELECT
   		department.DepName,
    	department.DepCode,
    	COUNT(DISTINCT itemstock_RFID.RfidCode) AS qty
		FROM damagenh
		INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
		INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
		INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
		INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
	    INNER JOIN itemstock_RFID ON SUBSTRING_INDEX(damagenh_detail_round.RFID, '#', 1) = SUBSTRING_INDEX(itemstock_RFID.RfidCode, '#', 1)

		WHERE   
    		damagenh_detail_round.DepCode != ''
            AND DATE(damagenh.DocDate) BETWEEN '" . $startDate . "' AND '" . $endDate . "'
            AND damagenh.IsStatus = 1
        GROUP BY
            department.DepName,
            department.DepCode
        ORDER BY department.DepName ASC
            ");

        // dd($mainItems);

        foreach ($mainItems as $item) {
            // Main row
            $rows[] = [$count++, $item->DepName, $item->qty, ''];
            $rowPointer++;
            $docDates = DB::select("
                    SELECT 
                    DISTINCT DATE(damagenh.DocDate) AS DocDate, 
                    COUNT(itemstock_RFID.RfidCode) AS qty
                    FROM 
                        damagenh
                    INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
                    INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
                    INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
                    INNER JOIN itemstock_RFID ON SUBSTRING_INDEX(damagenh_detail_round.RFID, '#', 1) = SUBSTRING_INDEX(itemstock_RFID.RfidCode, '#', 1)

                    WHERE 
                        department.DepCode = '" . $item->DepCode . "'
                        -- AND damagenh.DepCode = 'BPHCENTER' 
                        AND DATE(damagenh.DocDate) BETWEEN '" . $startDate . "' AND '" . $endDate . "'
                    GROUP BY 
                        DATE(damagenh.DocDate)  -- เพิ่ม GROUP BY
                    ORDER BY 
                        damagenh.DocDate

            ");

            foreach ($docDates as $doc) {
                $rows[] = ['', 'วันที่ ' . date('d/m/Y', strtotime("+543 years", strtotime($doc->DocDate))), $doc->qty, ''];

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
                        INNER JOIN itemstock_RFID ON SUBSTRING_INDEX(damagenh_detail_round.RFID, '#', 1) = SUBSTRING_INDEX(itemstock_RFID.RfidCode, '#', 1)

                    WHERE
                        department.DepCode = '" . $item->DepCode . "'
                        -- AND damagenh.DepCode = 'BPHCENTER' 
                        AND DATE(damagenh.DocDate) = '" . $doc->DocDate . "'
                    GROUP BY item.ItemName
                    ORDER BY item.ItemName
                    ");

                foreach ($subItems as $sub) {
                    $rows[] = ['', $sub->ItemName, $sub->qty];
                    $this->itemRows[] = $rowPointer++; // สำหรับ outline level 2
                }
            }
            $rows[] = ['', '', '', ''];
            $rowPointer++;
        }

        // dd($rows);

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
