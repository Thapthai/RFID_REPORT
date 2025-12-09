<table>
    <tr>
        <td></td>
        <td></td>
        <td colspan="5">วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ. {{ $currentYear + 543 }}
        </td>

    </tr>
    <tr>
        <th colspan="7">รายงานสต๊อกคงคลัง</th>
    </tr>
    <tr>
        <td style="border: 1px solid black">ลำดับ</td>
        <td style="border: 1px solid black">รายการ</td>
        <td style="border: 1px solid black">สต๊อกทั้งหมด</td>
        <td style="border: 1px solid black">อยู่ในโรงซัก</td>
        <td style="border: 1px solid black">อยู่ในห้องรับผ้าสะอาด </td>
        <td style="border: 1px solid black">อยู่ในแผนก </td>
        <td style="border: 1px solid black">ผ้าชำรุด
        </td>
    </tr>

    @php

        $i = 1;
    @endphp
    
    @foreach ($data as $item)
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td style="border: 1px solid black;">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->cntAll }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->cntDirty }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->cntClean }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->cntSticker }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->cntDam }}</td>
        </tr>
    @endforeach
</table>
