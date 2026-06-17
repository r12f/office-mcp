# office-ctl excel

Excel add-in entry point, host initialization, capability probing, and the
Excel v1 `excel.*` command implementations live here.

The task pane registers one workbook session with the daemon and advertises the
current v1 tool surface:

- `excel.read_range`
- `excel.write_range`
- `excel.add_sheet`
- `excel.set_formula`
- `excel.format_range`
- `excel.create_table`
- `excel.create_chart`

Run local checks from this directory:

```powershell
npm run check
```
