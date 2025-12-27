<?php

namespace App\Http\Controllers;

use App\Exports\ReportReportRoundFactoryXlsxExport;
use App\Exports\ReportDamageDailyMultiSheetXlsxExport;
use App\Exports\ReportDamageMultiSheetXlsxExport;
use App\Exports\ReportDamageSelectTypeMultiSheetXlsxExport;
use Maatwebsite\Excel\Facades\Excel;
use App\Exports\ReportStockBalanceXlsxExport;
use App\Exports\ReportStockMultiSheetXlsxExport;
use App\Exports\ReportUseLinenMultiSheetXlsxExport;
use App\Models\Departments;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class ExportExcelController extends Controller
{
    public function report_stock_balance_xlsx(Request $request)
    {

        $Date =  $request->query('Date');
        $Dep =  $request->query('Dep');

        // dd($Date);

        $results = DB::select(DB::raw("
        SELECT
            item.ItemCode,
            item.ItemName,
            COUNT(itf.ItemCode) AS TotalNum,
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.LastDocNo != 'CNPOS2000-00001' AND its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode) AS TotalUse,
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.LastDocNo != 'CNPOS2000-00001' AND its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF( DATE( NOW()), its.LastDocDate )) <= 30) AS 'Useloop',
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 1 ) AS TotalCancel,
            
            #เดือนก่อนหน้า
            (SELECT SUM(stock_balance.`stock<dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')  AND stock_balance.DepCode = '{$Dep}') AS TotalBefore1,
            (SELECT SUM(stock_balance.`stock>dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')  AND stock_balance.DepCode = '{$Dep}') AS TotalBefore2,
            (SELECT SUM(stock_balance.stockdamage) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')  AND stock_balance.DepCode = '{$Dep}') AS TotalBefore3,
            
            #เดือนที่เลือก
            (SELECT SUM(stock_balance.`stock<dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}'  AND stock_balance.DepCode = '{$Dep}') AS TotalSelect1,
            (SELECT SUM(stock_balance.`stock>dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}'  AND stock_balance.DepCode = '{$Dep}') AS TotalSelect2,
            (SELECT SUM(stock_balance.stockdamage) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}'  AND stock_balance.DepCode = '{$Dep}') AS TotalSelect3,
            
            #เดือนปัจจุบัน
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF( DATE( NOW()), its.LastDocDate )) < 7 AND   DATE_FORMAT(its.LastDocDate, '%m-%Y') = '{$Date}' AND its.LastDocNo != 'CNPOS2000-00001'  AND its.DepCode = '{$Dep}') AS TotalSelectNow1,
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF( DATE( NOW()), its.LastDocDate )) >= 7 AND its.LastDocNo != 'CNPOS2000-00001'  AND its.DepCode = '{$Dep}') AS TotalSelectNow2,
            (SELECT Count(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 1 AND DATE_FORMAT(its.LastDocDate, '%m-%Y') = '{$Date}' AND its.LastDocNo != 'CNPOS2000-00001'  AND its.DepCode = '{$Dep}') AS TotalSelectNow3,DATE_FORMAT(NOW(), '%m-%Y') As Date
            
    
        FROM
            itemstock_RFID itf
            INNER JOIN item ON item.ItemCode = itf.ItemCode
        WHERE
            itf.HptCode = 'BPH' 
        GROUP BY item.ItemCode, item.ItemName, itf.ItemCode, itf.HptCode

        ORDER BY
            item.ItemName ASC
    "));

        if ((string)$Dep === '0') {
            $results = DB::select(DB::raw("
            SELECT
                item.ItemCode,
                item.ItemName,
                COUNT(itf.ItemCode) AS TotalNum,
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.LastDocNo != 'CNPOS2000-00001' AND its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode) AS TotalUse,
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.LastDocNo != 'CNPOS2000-00001' AND its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF(DATE(NOW()), its.LastDocDate)) <= 30) AS Useloop,
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 1) AS TotalCancel,
                
                #เดือนก่อนหน้า
                (SELECT SUM(stock_balance.`stock<dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')) AS TotalBefore1,
                (SELECT SUM(stock_balance.`stock>dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')) AS TotalBefore2,
                (SELECT SUM(stock_balance.stockdamage) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = DATE_FORMAT(STR_TO_DATE(CONCAT('{$Date}', '-01'), '%m-%Y-%d') - INTERVAL 1 MONTH, '%m-%Y')) AS TotalBefore3,
        
                #เดือนที่เลือก
                (SELECT SUM(stock_balance.`stock<dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}') AS TotalSelect1,
                (SELECT SUM(stock_balance.`stock>dayfig`) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}') AS TotalSelect2,
                (SELECT SUM(stock_balance.stockdamage) FROM stock_balance WHERE stock_balance.ItemCode = itf.ItemCode AND stock_balance.HptCode = itf.HptCode AND DATE_FORMAT(stock_balance.DocDate, '%m-%Y') = '{$Date}') AS TotalSelect3,
        
                #เดือนปัจจุบัน
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF(DATE(NOW()), its.LastDocDate)) < 7 AND DATE_FORMAT(its.LastDocDate, '%m-%Y') = '{$Date}'  AND its.LastDocNo != 'CNPOS2000-00001') AS TotalSelectNow1,
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 0 AND (DATEDIFF(DATE(NOW()), its.LastDocDate)) >= 7 AND its.LastDocNo != 'CNPOS2000-00001') AS TotalSelectNow2,
                (SELECT COUNT(its.ItemCode) FROM itemstock_RFID its WHERE its.ItemCode = itf.ItemCode AND its.HptCode = itf.HptCode AND its.isCancel = 1 AND DATE_FORMAT(its.LastDocDate, '%m-%Y') = '{$Date}' AND its.LastDocNo != 'CNPOS2000-00001') AS TotalSelectNow3,
                DATE_FORMAT(NOW(), '%m-%Y') AS Date
        
            FROM
                itemstock_RFID itf
                INNER JOIN item ON item.ItemCode = itf.ItemCode
            WHERE
                itf.HptCode = 'BPH' 
            GROUP BY item.ItemCode, item.ItemName, itf.ItemCode, itf.HptCode
    
            ORDER BY
                item.ItemName ASC
        "));
        }

        // $data = response()->json($results);
        // return $data;

        return Excel::download(new ReportStockBalanceXlsxExport($results, $Date), 'stock_balance ' . $Date . '.xlsx');
    }

    public function report_stock_xlsx(Request $request)
    {

        $HptCode = 'BPH';

        if ($request->query('HptCode')) {
            $HptCode =  $request->query('HptCode');
        }
        // dd('stock');

        return Excel::download(new ReportStockMultiSheetXlsxExport($HptCode), 'stock ' . $HptCode . '.xlsx');
    }

    public function report_use_linen_xlsx(Request $request)
    {
        $year = date('Y');
        $month = date('m');

        if ($request->query('Month')) {
            $month = $request->query('Month');
            $month = explode("-", $month);
            $year = $month[1];
            $month = $month[0];
        }

        $daysInMonth = cal_days_in_month(CAL_GREGORIAN, $month, $year);

        $allDates = [];
        for ($day = 1; $day <= $daysInMonth; $day++) {
            $allDates[] = sprintf("%04d-%02d-%02d", $year, $month, $day);
        }

        // เตรียมข้อมูล department
        $depCode = $request->query('Dep');
        if ((string)$depCode === '0') {
            $departments = Departments::where('HptCode', 'BPH')
                ->where('IsActive', 1)
                ->orderBy('DepName', 'ASC')
                ->pluck('DepName', 'DepCode');
        } else {
            $departments = Departments::where('IsActive', 1)
                ->where('DepCode', $depCode)
                ->orderBy('DepName', 'ASC')
                ->pluck('DepName', 'DepCode');
        }

        // ดึงข้อมูล dirty, shelfcount, clean ของทั้งเดือน
        $dirtyAll = DB::table('dirty')
            ->select(DB::raw("DATE(dirty.DocDate) as date"), 'dirty_detail.DepCode', DB::raw("SUM(dirty_detail.Qty_item) as count"))
            ->leftJoin('dirty_detail', 'dirty_detail.DocNo', '=', 'dirty.DocNo')
            ->where('dirty.IsStatus', 1)
            // ->where('dirty.HptCode', 'BPH')
            ->whereMonth('dirty.DocDate', $month)
            ->whereYear('dirty.DocDate', $year);


        $shelfcountAll = DB::table('shelfcount')
            ->select(DB::raw("DATE(shelfcount.DocDate) as date"), 'shelfcount.DepCode', DB::raw("SUM(shelfcount_detail.TotalQty) as count"))
            ->leftJoin('shelfcount_detail', 'shelfcount_detail.DocNo', '=', 'shelfcount.DocNo')
            ->whereIn('shelfcount.IsStatus', [3, 4])
            // ->where('shelfcount.SiteCode', 'BPH')
            ->whereMonth('shelfcount.DocDate', $month)
            ->whereYear('shelfcount.DocDate', $year);

        $cleanAll = DB::table('clean')
            ->select(DB::raw("DATE(clean.DocDate) as date"), 'clean.DepCode', DB::raw("SUM(clean_detail.Qty) as count"))
            ->leftJoin('clean_detail', 'clean_detail.DocNo', '=', 'clean.DocNo')
            ->where('clean.IsStatus', 1)
            ->where('clean.HptCode', 'BPH')
            ->whereMonth('clean.DocDate', $month)
            ->whereYear('clean.DocDate', $year);



        if ((string)$depCode !== '0') {
            $dirtyAll = $dirtyAll->where('dirty_detail.DepCode', $depCode);
            $shelfcountAll = $shelfcountAll->where('shelfcount.DepCode', $depCode);
            $cleanAll = $cleanAll->where('clean.DepCode', $depCode);
        }

        $dirtyAll = $dirtyAll
            ->groupBy('dirty_detail.DepCode', DB::raw("DATE(dirty.DocDate)"))
            ->get()
            ->groupBy(function ($row) {
                return $row->DepCode . '|' . $row->date;
            });

        $shelfcountAll = $shelfcountAll
            ->groupBy('shelfcount.DepCode', DB::raw("DATE(shelfcount.DocDate)"))
            ->get()
            ->groupBy(function ($row) {
                return $row->DepCode . '|' . $row->date;
            });

        $cleanAll = $cleanAll
            ->groupBy('clean.DepCode', DB::raw("DATE(clean.DocDate)"))
            ->get()
            ->groupBy(function ($row) {
                return $row->DepCode . '|' . $row->date;
            });


        return Excel::download(new ReportUseLinenMultiSheetXlsxExport(
            $departments,
            $depCode,
            $allDates,
            $dirtyAll,
            $cleanAll,
            $shelfcountAll
        ), 'report_user_linen ' . $month . '-' . $year . '.xlsx');
    }


    public function report_damage_xlsx(Request $request)
    {

        $HptCode = 'BPH';

        if ($request->query('HptCode')) {
            $HptCode =  $request->query('HptCode');
        }
        $date = date('d-m-Y H:i:s');

        return Excel::download(new ReportDamageMultiSheetXlsxExport($HptCode), 'Report Damage (' . $date . ').xlsx');
    }

    public function report_damage_select_type_xlsx(Request $request)
    {

        $HptCode = 'BPH';
        $startDate = date('Y-m-d');
        $endDate = date('Y-m-d');

        $Type = 'Daily';
        $TypeTH = 'รายวัน';

 
        if ($request->query('hospital')) {
            $HptCode =  $request->query('hospital');
        }

        if ($request->query('search')) {
            // ?hospital=BPH&search=month&sDate=/undefined/NaN&eDate=/undefined/NaN&sMonth=03&eMonth=03&sYear=2025
            if ((string)$request->query('search') === 'month') {
                $Type = 'Month';
                $TypeTH = 'รายงานประเภท รายเดือน';
                $startMonth = $request->query('sMonth');
                $endMonth = $request->query('eMonth');
                $year = $request->query('sYear');
                $startDate = date('Y-m-d', strtotime('01-' . $startMonth . '-' . $year));

                $daysInMonth = date('t', strtotime('01-' . $endMonth . '-' . $year));
                $endDate = date('Y-m-d', strtotime($daysInMonth . '-' . $endMonth . '-' . $year));
            }

            // ?hospital=BPH&search=day&sDate=04/03/2025&eDate=04/03/2025&sMonth=&eMonth=&sYear=NaN
            if ((string)$request->query('search') === 'day') {
                $Type = 'Daily';
                $TypeTH = 'รายงานประเภท รายวัน';
                $sDate = $request->query('sDate');
                $eDate = $request->query('eDate');

                $sDate = str_replace('/', '-', $sDate);
                $eDate = str_replace('/', '-', $eDate);

                $startDate = date('Y-m-d', strtotime($sDate));
                $endDate = date('Y-m-d', strtotime($eDate));

                // ?hospital=BPH&search=day&sDate=01/03/2025&eDate=31/03/2025&sMonth=&eMonth=&sYear=NaN
                if ($sDate != $eDate) {
                    $Type = 'Between';
                    $TypeTH = 'รายงานระหว่าง ';
                }
            }
        }

        $date = date('d-m-Y H:i:s');

        return Excel::download(new ReportDamageSelectTypeMultiSheetXlsxExport($HptCode, $startDate, $endDate, $TypeTH), 'Damage' . $Type . 'Report(' . $date . ').xlsx');
    }

    public function report_report_round_factory_xlsx(Request $request)
    {
        // Increase memory limit for large datasets
        ini_set('memory_limit', '512M');
        set_time_limit(300); // 5 minutes timeout

        $HptCode = 'BPH';

        if ($request->query('HptCode')) {
            $HptCode = $request->query('HptCode');
        }

        $date = date('d-m-Y H:i:s');

        return Excel::download(new ReportReportRoundFactoryXlsxExport($HptCode), 'Report Round Factory(' . $date . ').xlsx');
    }
}
