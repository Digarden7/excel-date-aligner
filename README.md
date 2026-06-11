# Excel Date Aligner

pandas 按日期对齐合并 Excel 多工作表数据，支持缺失值前向填充。

## 功能说明

本工具用于将 Excel 文件中的多个工作表按「日期」列进行对齐合并，解决不同工作表中日期不一致、数据缺失的问题。

### 核心功能

- **多工作表合并**：读取 Excel 文件中的多个工作表，按「日期」列进行左连接合并
- **缺失值填充**：支持前向填充（ffill）和后向填充（bfill），补齐对齐后的缺失数据
- **自动命名输出**：默认在原文件名后添加 `_对齐后` 后缀生成输出文件

## 依赖

```
pandas
openpyxl
```

安装依赖：

```bash
pip install pandas openpyxl
```

## 使用方法

### 基本用法

```bash
python excel_date_aligner.py <输入文件>
```

### 完整参数

```bash
python excel_date_aligner.py <输入文件> [--output <输出文件>] [--method ffill|bfill]
```

### 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入 Excel 文件路径（必填） | - |
| `-o, --output` | 输出文件路径 | 自动生成（原文件名_对齐后.xlsx） |
| `-m, --method` | 缺失值填充方式 | `ffill`（前向填充） |

### 使用示例

```bash
# 基本用法：前向填充缺失值
python excel_date_aligner.py data.xlsx

# 指定输出文件名
python excel_date_aligner.py data.xlsx --output result.xlsx

# 使用后向填充
python excel_date_aligner.py data.xlsx --method bfill
```

## 输入文件要求

- 输入文件为 `.xlsx` 格式
- 每个工作表必须包含「日期」列
- 至少包含 2 个工作表才会执行合并

## 工作原理

1. 读取 Excel 文件中的所有工作表
2. 以第一个工作表为基准，按「日期」列进行左连接（left join）
3. 如有更多工作表，依次合并
4. 对合并后的缺失值进行填充（默认前向填充）
5. 输出合并结果到新的 Excel 文件

## License

MIT
