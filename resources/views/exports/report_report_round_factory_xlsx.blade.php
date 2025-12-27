<table>
    <tr>
        <td></td>
        <td></td>
        <td colspan="4">วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ. {{ $currentYear }}
        </td>
    </tr>
    <tr>
        <th colspan="6">รายงานรอบการซัก</th>
    </tr>
    <tr>
        <td style="border: 1px solid black">ลำดับ</td>
        <td style="border: 1px solid black">รายการ</td>
        <td style="border: 1px solid black">รหัส RFID</td>
        <td style="border: 1px solid black">วันที่สร้าง</td>
        <td style="border: 1px solid black">จำนวนรอบการซัก</td>
        <td style="border: 1px solid black">รหัส QR</td>
    </tr>

    @php
        $i = 1;
    @endphp
    @foreach ($data as $item)
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td style="border: 1px solid black">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black">{{ $item->RfidCode }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->Createdate }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->ReadCount }}</td>
            <td style="border: 1px solid black">{{ $item->QrCode }}</td>
        </tr>
    @endforeach
</table>

