<?php

namespace App\Exports\ReportUseLinen;

use Maatwebsite\Excel\Concerns\WithTitle;
use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithDrawings;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\WithEvents;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class ReportUseLinenSummarySheetXlsxExport implements FromView, WithTitle, WithDrawings, WithEvents
{
    public function title(): string
    {
        return 'รวมทั้งหมด';
    }

    public function view(): View
    {
        $date = date('Y-m');
        $year = date('Y');
        $month = date('m');

        $daysInMonth = cal_days_in_month(CAL_GREGORIAN, $month, $year);
        $DepName = 'รวมทั้งหมด';
        $summaryData = [];

        $dirtyData = DB::table('dirty')
            ->select(DB::raw("DATE(dirty.DocDate) as date"), DB::raw("SUM(dirty_detail.Qty_item) as count"))
            ->leftJoin('dirty_detail', 'dirty_detail.DocNo', '=', 'dirty.DocNo')
            ->where('dirty.IsStatus', 1)
            ->where('dirty.HptCode', 'BPH')
            ->whereMonth('dirty.DocDate', $month)
            ->whereYear('dirty.DocDate', $year)
            ->groupBy(DB::raw("DATE(dirty.DocDate)"))
            ->pluck('count', 'date');

        $shelfcountData = DB::table('shelfcount')
            ->select(DB::raw("DATE(shelfcount.DocDate) as date"), DB::raw("SUM(shelfcount_detail.TotalQty) as count"))
            ->leftJoin('shelfcount_detail', 'shelfcount_detail.DocNo', '=', 'shelfcount.DocNo')
            ->whereIn('shelfcount.IsStatus', [3, 4])
            ->where('shelfcount.SiteCode', 'BPH')
            ->whereMonth('shelfcount.DocDate', $month)
            ->whereYear('shelfcount.DocDate', $year)
            ->groupBy(DB::raw("DATE(shelfcount.DocDate)"))
            ->pluck('count', 'date');

        $cleanData = DB::table('clean')
            ->select(DB::raw("DATE(clean.DocDate) as date"), DB::raw("SUM(clean_detail.Qty) as count"))
            ->leftJoin('clean_detail', 'clean_detail.DocNo', '=', 'clean.DocNo')
            ->where('clean.IsStatus', 1)
            ->where('clean.HptCode', 'BPH')
            ->whereMonth('clean.DocDate', $month)
            ->whereYear('clean.DocDate', $year)
            ->groupBy(DB::raw("DATE(clean.DocDate)"))
            ->pluck('count', 'date');

        for ($day = 1; $day <= $daysInMonth; $day++) {
            $dateFormatted = sprintf("%04d-%02d-%02d", $year, $month, $day);
            $summaryData[] = [
                'date' => $dateFormatted,
                'dirtyLinen' => $dirtyData[$dateFormatted] ?? 0,
                'cleanLinen' => $cleanData[$dateFormatted] ?? 0,
                'Shelfcount' => $shelfcountData[$dateFormatted] ?? 0,
            ];
        }

        $reportDate =  date('m-Y');
        $thaiMonths = config('myconfig.thai_months');

        // แปลงเป็น timestamp
        $reportTimestamp = strtotime('01-' . $reportDate);
        $currentMonthNum = date('m', $reportTimestamp);

        $currentMonthName = $thaiMonths[$currentMonthNum];

        $currentYear = date('Y', $reportTimestamp);

        return view('exports.reportUseLinen_xlsx.reportUseLinenSheet', compact(
            'summaryData',
            'DepName',
            'currentMonthName',
            'currentYear'
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

                $sheet->getStyle('B1')->applyFromArray([
                    'font' => [
                        'name' => 'Angsana New',
                        'size' => 28,
                        'bold' => true
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
                    'A' => 25,
                    'B' => 25,
                    'C' => 25,
                    'D' => 25,

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

                $sheet->getStyle('B1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle('B1')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

                $sheet->getStyle("A3:D{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle("A3:D{$highestRow}")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
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
        $drawing->setOffsetX(30);
        $drawing->setOffsetY(10);

        return [$drawing];
    }
}
