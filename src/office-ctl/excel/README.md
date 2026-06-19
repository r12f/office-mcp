# office-ctl excel

Excel add-in entry point, host initialization, capability probing, and the
Excel v1 `excel.*` command implementations live here.

The task pane registers one workbook session with the daemon and currently
advertises the implemented v1 tool surface:

- `excel.get_workbook_info`
- `excel.list_sheets`
- `excel.add_sheet`
- `excel.update_sheet`
- `excel.delete_sheet`
- `excel.get_used_range`
- `excel.read_range`
- `excel.write_range`
- `excel.clear_range`
- `excel.find_replace_cells`
- `excel.set_formula`
- `excel.format_range`
- `excel.create_table`
- `excel.create_chart`

The target core Excel backlog is the refined about-20-tool catalog documented in
`doc/spec/04-excel-capabilities.md` and tracked in `doc/spec/08-roadmap.md` under
M7.1. It is based on the workbook, worksheet, range, table, chart, and PivotTable
workflow from the Microsoft Excel add-in API docs.

Run local checks from this directory:

```powershell
npm run check
```
