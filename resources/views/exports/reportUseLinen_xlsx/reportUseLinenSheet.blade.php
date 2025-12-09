<table>
    <tr>
        <td></td>
        <th colspan="3">รายงานการใช้งานผ้าของหน่วยงาน</th>

    </tr>
    <tr>
        <td colspan="2">แผนก : {{ $DepName }}</td>
        <td colspan="2" style="text-align: right">วันที่พิมพ์รายงาน {{ date('d') }} {{ $currentMonthName }} พ.ศ.
            {{ $currentYear + 543 }}
        </td>
    </tr>
    <tr>
        <td style="border: 1px solid black;background-color: #8fdaff">วันที่เอกสาร</td>
        <td style="border: 1px solid black;background-color: #8fdaff">ผ้าเปื้อนส่งซัก</td>
        <td style="border: 1px solid black;background-color: #8fdaff">รับเข้าผ้าสะอาด</td>
        <td style="border: 1px solid black;background-color: #8fdaff">shelf count</td>
    </tr>

    @foreach ($summaryData as $day)
        <tr>
            <td style="border: 1px solid black;">{{ $day['date'] }}</td>
            <td style="border: 1px solid black;">{{ $day['dirtyLinen'] }}</td>
            <td style="border: 1px solid black;"> {{ $day['cleanLinen'] }}</td>
            <td style="border: 1px solid black;">{{ $day['Shelfcount'] }}</td>
        </tr>
    @endforeach
    @php
        $totalDirty = collect($summaryData)->sum('dirtyLinen');
        $totalClean = collect($summaryData)->sum('cleanLinen');
        $totalShelfcount = collect($summaryData)->sum('Shelfcount');
    @endphp
    <tr>
        <td style="border: 1px solid black; text-align: center;background-color: #aefcd4"><strong>รวม</strong></td>
        <td style="border: 1px solid black;background-color: #aefcd4"><strong>{{ $totalDirty }}</strong></td>
        <td style="border: 1px solid black;background-color: #aefcd4"><strong>{{ $totalClean }}</strong></td>
        <td style="border: 1px solid black;background-color: #aefcd4"><strong>{{ $totalShelfcount }}</strong></td>
    </tr>
</table>
