<?php

namespace App\Exports;

use App\Exports\ReportDamageLinenSelectType\ReportDamageLinenRawSelectTypeSheetXlsxExport;
use App\Exports\ReportDamageLinenSelectType\ReportDamageLinenSelectTypeSheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Illuminate\Support\Facades\DB;


class ReportDamageSelectTypeMultiSheetXlsxExport implements WithMultipleSheets
{
    protected $HptCode;
    protected $startDate;
    protected $endDate;
    protected $typeTH;
    public function __construct(
        $HptCode,
        $startDate,
        $endDate,
        $typeTH
    ) {
        $this->HptCode = $HptCode;
        $this->startDate = $startDate;
        $this->endDate = $endDate;
        $this->typeTH = $typeTH;
    }
    public function sheets(): array
    {
        $HptCode = $this->HptCode;
        $startDate = $this->startDate;
        $endDate = $this->endDate;
        $typeTH = $this->typeTH;

        $thaiMonths = config('myconfig.thai_months');
        $startDate = $this->startDate;
        $startDateDay = date('d',  strtotime($startDate));
        $startDateMonthNameTH = $thaiMonths[date('m',  strtotime($startDate))];
        $startDateYear =  date('Y', strtotime("+543 years", strtotime($startDate)));

        $endDate = $this->endDate;
        $endDateDay = date('d',  strtotime($endDate));
        $endDateMonthNameTH = $thaiMonths[date('m',  strtotime($endDate))];
        $endDateYear =  date('Y', strtotime("+543 years", strtotime($endDate)));

        $typeTopic = $typeTH . ' วันที่ ' . $startDateDay . ' ' . $startDateMonthNameTH . ' ' . $startDateYear .
            ($startDate != $endDate ? ' ถึง ' . $endDateDay . ' ' . $endDateMonthNameTH . ' ' . $endDateYear : '');

        // ดึงข้อมูลทั้งหมดครั้งเดียว
        $allData = DB::select("
            SELECT
                department.DepName,
                department.DepCode,
                item.ItemName,
                damagenh_detail_round.RFID,
                damagenh_detail_round.QrCode,
                damagenh.DocDate AS DocDate,
                itemstock_RFID.ReadCount AS ReadCount,
                itemstock_RFID.RfidCode
            FROM damagenh
            INNER JOIN damagenh_detail ON damagenh.DocNo = damagenh_detail.DocNo
            INNER JOIN damagenh_detail_round ON damagenh_detail.Id = damagenh_detail_round.RowID
            INNER JOIN item ON damagenh_detail_round.ItemCode = item.ItemCode
            INNER JOIN department ON damagenh_detail_round.DepCode = department.DepCode
            INNER JOIN itemstock_RFID ON SUBSTRING_INDEX(damagenh_detail_round.RFID, '#', 1) = SUBSTRING_INDEX(itemstock_RFID.RfidCode, '#', 1) 
            WHERE DATE(damagenh.DocDate) BETWEEN '" . $startDate . "' AND '" . $endDate . "'
                AND damagenh_detail_round.DepCode != ''
                AND damagenh.IsStatus = 1
            ORDER BY department.DepName, damagenh.DocDate, item.ItemName ASC
        ");

        // จัดกลุ่มข้อมูลใน PHP (Cleaning Data)
        $groupedData = [];
        $rawData = []; // เก็บข้อมูล raw สำหรับ sheet ที่ 2

        foreach ($allData as $row) {
            $depCode = $row->DepCode;
            $depName = $row->DepName;
            $docDate = date('Y-m-d', strtotime($row->DocDate));
            $itemName = $row->ItemName;
            $rfidCode = $row->RfidCode;

            // สร้าง structure: Department -> DocDate -> ItemName -> RfidCodes[]
            if (!isset($groupedData[$depCode])) {
                $groupedData[$depCode] = [
                    'DepName' => $depName,
                    'DepCode' => $depCode,
                    'dates' => [],
                    'rfids' => [] // สำหรับนับจำนวนรวมของ Department
                ];
            }

            if (!isset($groupedData[$depCode]['dates'][$docDate])) {
                $groupedData[$depCode]['dates'][$docDate] = [
                    'items' => [],
                    'rfids' => [] // สำหรับนับจำนวนรวมของวันที่
                ];
            }

            if (!isset($groupedData[$depCode]['dates'][$docDate]['items'][$itemName])) {
                $groupedData[$depCode]['dates'][$docDate]['items'][$itemName] = [];
            }

            // เก็บ RfidCode ที่ไม่ซ้ำกัน
            $groupedData[$depCode]['dates'][$docDate]['items'][$itemName][$rfidCode] = true;
            $groupedData[$depCode]['dates'][$docDate]['rfids'][$rfidCode] = true;
            $groupedData[$depCode]['rfids'][$rfidCode] = true;

            // เก็บข้อมูล raw สำหรับ sheet ที่ 2
            $rawData[] = $row;
        }
 
        return [
            new ReportDamageLinenSelectTypeSheetXlsxExport($HptCode, $startDate, $endDate, $typeTopic, $groupedData),
            new ReportDamageLinenRawSelectTypeSheetXlsxExport($HptCode, $startDate, $endDate, $typeTopic, $rawData),
        ];
    }
}
