# source_data.cls

## merge_worksheet_in_multiple_workbooks

- Merge worksheets in multiple workbooks into one workbook.

| Parameter | Type | Description |
| --- | --- | --- |
| workbook_path_list | Variant | The list of the paths of the workbooks to be merged |
| worksheet_name | String | The name of the worksheet to be merged |
| output_workbook | Workbook | The workbook to be output |
| output_worksheet_name_list | Variant | The list of the names of the worksheets to be output |
| output_worksheet_name_prefix | String | The prefix of the names of the worksheets to be output |
| output_worksheet_name_suffix | String | The suffix of the names of the worksheets to be output |

## reformulate_worksheets

- Re-formulate every cell in column C to column F

| Parameter | Type | Description |
| --- | --- | --- |
| wb | Workbook | The workbook to be reformulated |
| ws_name | String | The name of the worksheet to be reformulated |
| avg_or_sum | String | The type of the reformulation, "avg" or "sum" |
