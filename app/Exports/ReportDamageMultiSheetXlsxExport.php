<?php

namespace App\Exports;

use App\Exports\ReportDamageLinen\ReportDamageLinenRawSheetXlsxExport;
use App\Exports\ReportDamageLinen\ReportDamageLinenSheetXlsxExport;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Illuminate\Support\Facades\DB;


class ReportDamageMultiSheetXlsxExport implements WithMultipleSheets
{
    protected $HptCode;

    public function __construct(
            $HptCode
    ) {
        $this->HptCode = $HptCode;
    }
    public function sheets(): array
    {
        $HptCode = $this->HptCode;
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
            WHERE damagenh_detail_round.DepCode != ''
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
            new ReportDamageLinenSheetXlsxExport($HptCode, $groupedData),
            new ReportDamageLinenRawSheetXlsxExport($HptCode, $rawData),
        ];
    }
}
