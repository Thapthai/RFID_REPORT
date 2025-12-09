{{-- <table>
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
        <td colspan="3" style="border: 1px solid black">รายการ</td>
        <td style="border: 1px solid black">จำนวน</td>
    </tr>

    @php

        $i = 1;
    @endphp
    @foreach ($data as $item)
        <tr>
            <td style="border: 1px solid black;text-align: center">{{ $i++ }}</td>
            <td colspan="3" style="border: 1px solid black">{{ $item->ItemName }}</td>
            <td style="border: 1px solid black;text-align: center">{{ $item->Qty }}</td>

        </tr>
    @endforeach
</table> --}}


<table>
    <tr>
        <td colspan="3">วันที่พิมพ์รายงาน{{ $printdate }}</td>
    </tr>
    <tr>
        <td colspan="3" style="text-align: center;"><strong>รายงานสต๊อกคงคลัง</strong></td>
    </tr>
    <tr>
        <th>ลำดับ</th>
        <th>รายการ</th>
        <th>จำนวน</th>
    </tr>
    @php $count = 1; @endphp
    @foreach ($data as $row)
        <tr>
            <td>{{ $count++ }}</td>
            <td>{{ $row->ItemName }}</td>
            <td>{{ $row->Qty }}</td>
        </tr>
        @foreach ($subData[$row->ItemCode] as $sub)
            <tr>
                <td></td>
                <td>
                    @if ($sub->StatusDepartment == 1)
                        {{ $sub->DepName }}
                    @elseif ($sub->StatusLaundry == 1)
                        {{ $sub->FacName }}
                    @elseif ($sub->StatusLinenClean == 1)
                        {{ $sub->DepName }}
                    @endif
                </td>
                <td>{{ $sub->Qty }}</td>
            </tr>
        @endforeach
    @endforeach
</table>
