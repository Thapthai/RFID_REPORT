<table>
    <tr>
        <td></td>
        <td></td>
        <td colspan="3">วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ. {{ $currentYear + 543 }}
        </td>

    </tr>
    <tr>
        <th colspan="5">รายงานสต๊อกคงคลัง</th>
    </tr>
    <tr>
        <td style="border: 1px solid black">ลำดับ</td>
        <td style="border: 1px solid black">รายการ</td>
        <td style="border: 1px solid black">แหล่งที่มา</td>
        <td style="border: 1px solid black">ที่อยู่ปัจจุบัน</td>
        <td style="border: 1px solid black">จำนวน</td>
    </tr>

    @php

        $i = 1;
    @endphp
    @foreach ($data as $item)
        @php
            $statustext = '';
            if ($item->StatusDepartment == 1) {
                $statustext = 'Department';
            }
            if ($item->StatusLaundry == 1) {
                $statustext = 'Laundry';
            }
            if ($item->StatusLinenClean == 1) {
                $statustext = 'Center';
            }
        @endphp
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td style="border: 1px solid black">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black">{{ $statustext }}</td>
            <td style="border: 1px solid black">{{ $item->DepName }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->Qty }}</td>

        </tr>
    @endforeach
</table>
