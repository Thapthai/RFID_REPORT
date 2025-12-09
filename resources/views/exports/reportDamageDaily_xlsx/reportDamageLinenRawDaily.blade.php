<table>
    <tr>
        <td></td>
        <td></td>
        <td colspan="5">วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ. {{ $currentYear + 543 }}
        </td>

    </tr>
    <tr>
        <th colspan="7">รายงานผ้าชำรุด</th>
    </tr>
    <tr>
        <td style="border: 1px solid black"> ลำดับ</td>
        <td style="border: 1px solid black">แผนก</td>
        <td style="border: 1px solid black">รายการ</td>
        <td style="border: 1px solid black">RFID</td>
        <td style="border: 1px solid black">QRCODE</td>
        <td style="border: 1px solid black">วันที่ชำรุด</td>
        <td style="border: 1px solid black">จำนวนรอบการซัก</td>
    </tr>

    @php

        $i = 1;
    @endphp
    @foreach ($data as $item)
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td style="border: 1px solid black">{{ $item->DepName }}</td>
            <td style="border: 1px solid black">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black">{{ $item->RFID }}</td>
            <td style="border: 1px solid black">{{ $item->QrCode }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->DocDate }}</td>
            <td style="border: 1px solid black;text-align: center"> {{ number_format($item->ReadCount, 0) }}</td>
        </tr>
    @endforeach

</table>
