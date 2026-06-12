# AttendanceGenerator

C# / .NET 8 / WinForms / ClosedXML 考勤统计表生成器。

## 功能

- 选择考勤报表、加班表两个业务文件。
- 输入月份，例如 `2026-05`。
- 在界面上勾选本月法定节假日。
- 使用固定模板 `Config/考勤统计模板.xlsx`。
- 基于模板文件复制生成新 Excel，不重新画表。
- 自动识别考勤主表中的姓名列、部门列、日期列、统计列。
- 自动识别加班补贴表中的日期矩阵。
- 生成 `YYYY年MM月考勤统计表_生成版.xlsx`。
- 无法识别的数据写入 `异常明细` Sheet。
- 运行日志写入 `Logs` 目录。

## Excel 识别规则

程序会优先从考勤报表的 `月度汇总` Sheet 读取每日考勤结果区，识别 `姓名`、`部门`、`考勤结果` 等表头；如果识别失败，会退回读取 `每日统计` Sheet。

加班表会自动识别审批导出明细中的：

- `发起人姓名` 或 `加班人`
- `开始时间` 或 `加班时间`
- `时长`
- `加班类型`
- `审批结果`

审批结果包含 `拒绝` 或 `撤销` 的记录会被跳过。

模板固定为程序目录下的 `Config/考勤统计模板.xlsx`。程序会按该模板复制，然后在复制件中写入数据。程序不会重建 Sheet，因此模板中的 Sheet 名称、合并单元格、行高、列宽、字体、边框、填充色、对齐方式、公式和打印设置会尽量保留。

如需更换模板，替换发布目录或源码目录中的同名文件即可：

```text
Config\考勤统计模板.xlsx
```

## 配置

界面会根据月份列出当月所有日期，默认不勾选；用户按实际制度自行勾选或取消。生成时以界面勾选结果为准。

`Config/holiday.json` 用于无 UI 调用或默认规则补充：

```json
{
  "mode": "AutoChineseStatutory",
  "extraHolidays": [],
  "workdays": []
}
```

- `extraHolidays`：额外增加为法定的日期。
- `workdays`：从自动法定节假日中排除的日期。

`Config/mapping.json` 配置状态映射。默认包含：

- `正常出勤` / `普通出勤` / `正常` => `公司出勤`
- `休息` => `休`
- `法定节假日` => `法定`
- `事假` / `调休` / `年假` / `病假` / `婚假` / `产假` / `丧假` => `类型XH`
- `旷工` / `迟到` / `早退` / `缺卡` => 原状态

## 运行

```powershell
dotnet restore
dotnet run --project .\AttendanceGenerator.csproj
```


## 重新生成/发布说明

仓库不再提交 `bin/`、`obj/` 或 `.vs/` 下的本地构建产物。拉取代码后请重新编译或发布，不要运行旧的 `bin\Release` 或 `publish` 目录中的 exe：

```powershell
dotnet restore .\AttendanceGenerator.csproj
dotnet publish .\AttendanceGenerator.csproj -c Release -f net8.0-windows -r win-x64 --self-contained false -p:PublishSingleFile=true
```

如果需要复核 5 月输出，也可以在 Windows/.NET SDK 环境中使用命令行入口：

```powershell
dotnet run --project .\AttendanceGenerator.csproj -- generate .\考勤报表.xlsx .\加班表.xlsx 2026-05 .\out 2026-05-01,2026-05-02 .\Config\考勤统计模板.xlsx
dotnet run --project .\AttendanceGenerator.csproj -- compare .\out\2026年05月考勤统计表_生成版.xlsx .\5月手工做的.xlsx
```

## 打包为 exe

框架依赖发布：

```powershell
dotnet publish .\AttendanceGenerator.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
```

自包含发布：

```powershell
dotnet publish .\AttendanceGenerator.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

发布文件位于：

```text
bin\Release\net8.0-windows\win-x64\publish\
```

运行发布后的 `AttendanceGenerator.exe` 即可。
