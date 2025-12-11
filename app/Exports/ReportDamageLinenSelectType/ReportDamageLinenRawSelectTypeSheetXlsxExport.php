<?php

namespace App\Exports\ReportDamageLinenSelectType;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Maatwebsite\Excel\Concerns\WithTitle;
use Illuminate\Support\Facades\DB;


class ReportDamageLinenRawSelectTypeSheetXlsxExport implements FromView, WithDrawings, WithEvents, WithTitle
{

    protected $HptCode;
    protected $startDate;
    protected $endDate;
    protected $typeTopic;


    public function __construct($HptCode, $startDate, $endDate, $typeTopic)

    {
        $this->HptCode = $HptCode;
        $this->startDate = $startDate;
        $this->endDate = $endDate;
        $this->typeTopic = $typeTopic;
    }

    public function title(): string
    {
        return   "รายการผ้าชำรุด ดิบ";
    }

    public function view(): View
    {
        $HptCode = $this->HptCode;
        $typeTopic = $this->typeTopic;
        $startDate = $this->startDate;
        $endDate = $this->endDate;

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

        // dd($startDate, $endDate);

        $data = DB::select("
            SELECT
                department.DepName,
                item.ItemName,
                damagenh_detail_round.RFID,
                damagenh_detail_round.QrCode,
                MIN(damagenh.DocDate) AS DocDate,
                MIN(itemstock_RFID.ReadCount) AS ReadCount
            FROM damagenh
            INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
            INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
            INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
            INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
            INNER JOIN itemstock_RFID ON SUBSTRING_INDEX(damagenh_detail_round.RFID, '#', 1) = SUBSTRING_INDEX(itemstock_RFID.RfidCode, '#', 1) 
            WHERE DATE(damagenh.DocDate) BETWEEN '".$startDate."' AND '".$endDate."'
            GROUP BY damagenh_detail_round.RFID
            ORDER BY department.DepName, item.ItemName ASC;

        ");


        return view('exports.reportDamageSelectType_xlsx.reportDamageLinenRawSelectType', compact(
            'typeTopic',
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
                        'bold' => true,
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

                $highestRow = $sheet->getHighestRow(); // หาบรรทัดสุดท้ายที่มีข้อมูล
                $sheet->getStyle("A4:Z{$highestRow}")->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 16,
                    ],
                ]);

                $event->sheet->getDelegate()->getStyle('A4:G4')->applyFromArray([
                    'fill' => [
                        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => '8fdaff',
                        ],
                    ],
                ]);


                // ตั้งค่าความกว้างของคอลัมน์
                $columns = [
                    'A' => 5,
                    'B' => 45,
                    'C' => 25,
                    'D' => 35,
                    'E' => 20,
                    'F' => 20,
                    'G' => 20,
                ];
                foreach ($columns as $col => $width) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                // ปรับความสูงของแถว
                $rows = [
                    2 => 50,
                    3 => 30,

                ];

                foreach ($rows as $row => $height) {
                    $sheet->getRowDimension($row)->setRowHeight($height);
                }

                $sheet->getStyle('C1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $sheet->getStyle('C1')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

                $sheet->getStyle('A2:L3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A2:L3')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
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
