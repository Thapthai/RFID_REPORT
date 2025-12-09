<table>

    <tr>
        <td colspan="2"></td>
        <th colspan="10">รายงานสต๊อกคงคลัง</th>
    </tr>
    <tr>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td>วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ. {{ $currentYear + 543 }}</td>
    </tr>
    <tr>
        <td rowspan="2" style="border: 1px solid black">ลำดับ</td>
        <td rowspan="2" style="border: 1px solid black">รายการ</td>
        <td rowspan="2" style="border: 1px solid black; white-space: normal; word-wrap: break-word;">
            สต็อกทั้งหมดในระบบ
        </td>

        <td rowspan="2" style="border: 1px solid black; white-space: normal; word-wrap: break-word;">สต๊อกที่ถูกใช้งาน
        </td>
        <td rowspan="2" style="border: 1px solid black; white-space: normal; word-wrap: break-word;">
            สต๊อกหมุนเวียนในระบบ</td>
        <td rowspan="2" style="border: 1px solid black; white-space: normal; word-wrap: break-word;">ผ้าชำรุดในระบบ
        </td>
        <td colspan="3" style="border: 1px solid black">เดือน {{ $previousMonthName }} {{ $previousYear }}</td>
        <td colspan="3" style="border: 1px solid black">เดือน {{ $currentMonthName }} {{ $currentYear }}</td>
    </tr>
    <tr>
        <td style="border: 1px solid black">สต๊อกเคลื่อนไหว</td>
        <td style="border: 1px solid black">ไม่เคลื่อนไหว</td>
        <td style="border: 1px solid black"> ผ้าชำรุด</td>
        <td style="border: 1px solid black">สต๊อกเคลื่อนไหว</td>
        <td style="border: 1px solid black">ไม่เคลื่อนไหว</td>
        <td style="border: 1px solid black"> ผ้าชำรุด</td>
    </tr>
    @php
        $i = 1;
    @endphp
    @foreach ($data as $item)
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td style="border: 1px solid black">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black">{{ $item->TotalNum }}</td>
            <td style="border: 1px solid black">{{ $item->TotalUse }}</td>
            <td style="border: 1px solid black">{{ $item->Useloop }}</td>
            <td style="border: 1px solid black">{{ $item->TotalCancel }}</td>


            {{-- <td style="border: 1px solid black;background-color: #eff488">{{ $item->TotalBefore1 ?? '' }}</td>
            <td style="border: 1px solid black;background-color: #eff488">{{ $item->TotalBefore2 ?? '' }}</td>
            <td style="border: 1px solid black;background-color: #eff488">{{ $item->TotalBefore3 ?? '' }}</td>

            @if ($reportDate == date('m-Y'))
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelectNow1 ?? '' }}</td>
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelectNow2 ?? '' }}</td>
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelectNow3 ?? '' }}</td>
            @else
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelect1 ?? '' }}</td>
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelect2 ?? '' }}</td>
                <td style="border: 1px solid black;background-color: #a7f1bb">{{ $item->TotalSelect3 ?? '' }}</td>
            @endif --}}

            <td style="border: 1px solid black;background-color: #FFEB9C">
                {{ empty($item->TotalBefore1) ? '' : $item->TotalBefore1 }}</td>
            <td style="border: 1px solid black;background-color: #FFEB9C">
                {{ empty($item->TotalBefore2) ? '' : $item->TotalBefore2 }}</td>
            <td style="border: 1px solid black;background-color: #FFEB9C">
                {{ empty($item->TotalBefore3) ? '' : $item->TotalBefore3 }}</td>

            @if ($reportDate == date('m-Y'))
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelectNow1) ? '' : $item->TotalSelectNow1 }}</td>
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelectNow2) ? '' : $item->TotalSelectNow2 }}</td>
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelectNow3) ? '' : $item->TotalSelectNow3 }}</td>
            @else
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelect1) ? '' : $item->TotalSelect1 }}</td>
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelect2) ? '' : $item->TotalSelect2 }}</td>
                <td style="border: 1px solid black;background-color: #C6EFCD">
                    {{ empty($item->TotalSelect3) ? '' : $item->TotalSelect3 }}</td>
            @endif

        </tr>
    @endforeach
</table>
