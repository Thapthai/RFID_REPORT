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
use Illuminate\Support\Facades\DB;

class ReportReportRoundFactoryXlsxExport implements FromView, WithDrawings, WithEvents, WithTitle
{
    protected $HptCode;

    public function __construct($HptCode)
    {
        $this->HptCode = $HptCode;
    }

    public function title(): string
    {
        return "รายงานรอบการซัก";
    }

    public function view(): View
    {
        $HptCode = $this->HptCode;

        $reportDate = date('d-m-Y H:i:s');
        $thaiMonths = config('myconfig.thai_months');
        $currentMonthNum = date('m');
        $currentMonthName = $thaiMonths[$currentMonthNum] ?? date('F');
        $currentYear = date('Y') + 543;

        // Use query builder with chunking for better memory management
        // Process data in chunks to avoid memory exhaustion
        $data = [];
        $chunkSize = 1000;
        
        DB::table('itemstock_RFID')
            ->select([
                'item.ItemName',
                'itemstock_RFID.RfidCode',
                DB::raw("DATE_FORMAT(insertrfid.Createdate,'%d-%m-%Y') AS Createdate"),
                'itemstock_RFID.ReadCount',
                'itemstock_RFID.QrCode'
            ])
            ->join('insertrfid', 'itemstock_RFID.InsertrfidDocNo', '=', 'insertrfid.DocNo')
            ->join('item', 'itemstock_RFID.ItemCode', '=', 'item.ItemCode')
            ->where('itemstock_RFID.HptCode', $HptCode)
            ->orderBy('item.ItemName', 'ASC')
            ->orderBy('insertrfid.Createdate', 'ASC')
            ->orderBy('itemstock_RFID.QrCode', 'ASC')
            ->chunk($chunkSize, function ($chunk) use (&$data) {
                foreach ($chunk as $row) {
                    $data[] = $row;
                }
            });

        return view('exports.report_report_round_factory_xlsx', compact(
            'data',
            'reportDate',
            'currentMonthName',
            'currentYear'
        ));
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                $sheet->getPageSetup()->setScale(100);
                $sheet->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);
                $sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_PORTRAIT);

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

                $highestRow = $sheet->getHighestRow();
                $sheet->getStyle("A3:F{$highestRow}")->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 16,
                    ],
                ]);

                $columns = [
                    'A' => 5,
                    'B' => 35,
                    'C' => 20,
                    'D' => 15,
                    'E' => 15,
                    'F' => 20,
                ];
                foreach ($columns as $col => $width) {
                    $sheet->getColumnDimension($col)->setWidth($width);
                }

                $rows = [
                    2 => 50,
                ];

                foreach ($rows as $row => $height) {
                    $sheet->getRowDimension($row)->setRowHeight($height);
                }

                $sheet->getStyle('C1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $sheet->getStyle('C1')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

                $sheet->getStyle('A2:F3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('A2:F3')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

                $event->sheet->getDelegate()->getStyle('A3:F3')->applyFromArray([
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
