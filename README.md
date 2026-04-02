# Bidding Script

用于援外项目投标文件的自动化生成，当前核心能力是从 `project*.docx` 提取项目信息并生成报价表 Excel。

## 当前状态

- 已启用入口功能：
  - 读取 `project*.docx`
  - 解析项目基础信息、物资清单、服务需求
  - 生成总报价表 `投标报价表-<项目名>.xlsx`
- 已实现但当前入口未启用：
  - `directory.py`：创建投标文件目录结构
  - `content.py`：生成投标目录 Excel
  - `cover.py`：生成封面
  - `separate.py`：将报价表按指定工作表拆分为单独文件，并尽量固化公式结果

## 依赖环境

- Python 3.10+
- `python-docx`
- `openpyxl`

安装依赖：

```bash
pip install python-docx openpyxl
```

`separate.py` 如果要把公式结果稳定固化为数值，建议本机安装 Microsoft Excel。没有 Excel 时，仍可拆分，但依赖工作簿中已有的公式缓存值。

## 输入文件

将项目 Word 文件放在项目根目录，文件名需匹配：

- `project*.docx`

示例：

- `project-[Project Name].docx`

建议同目录只保留一个匹配的 Word 文件，入口脚本会使用最后匹配到的文件名。

## 运行方式

当前默认入口：

```bash
python main.py
```

当前 `main.py` 会执行：

1. 查找并读取 `project*.docx`
2. 构造 `Project`
3. 生成报价表 `投标报价表-<项目名>.xlsx`

## 输出结果

默认运行 `main.py` 后，会在项目根目录生成：

- `投标报价表-<项目名>.xlsx`

如果手动启用 `separate.py`，还会额外生成按工作表拆分的文件，例如：

- `1.投标报价总表.xlsx`
- `2.采购需求偏离表(物资部分).xlsx`
- `3.开标一览表.xlsx`

## 代码结构

- `main.py`：入口脚本，当前默认只生成报价表
- `project.py`：读取并解析 Word 项目信息
- `quotation.py`：生成完整报价表工作簿
- `separate.py`：拆分报价表，并用缓存值导出单表文件
- `directory.py`：创建投标文件目录结构
- `content.py`：生成投标目录 Excel
- `cover.py`：封面生成

## 报价表说明

- 报价表由 `quotation.py` 直接生成，不依赖 Excel 模板。
- 工作簿内大量金额和汇总单元格使用公式。
- `wb.calculation.fullCalcOnLoad = True` 已开启，Excel 打开文件后会触发重算。
- 新版 Excel 可能会把部分公式显示为 `=@INDEX(...)`，这是动态数组兼容行为，通常不影响结果。

## 拆分功能说明

`separate.py` 的主要逻辑：

- 查找 `投标报价表-*.xlsx`
- 先尝试调用本机 Excel 打开并保存一次，以刷新公式缓存值
- 使用 `data_only=True` 重新读取工作簿
- 按工作表名匹配 `^[0-9]{1,2}\.` 的工作表分别导出为单独文件

如果需要在入口中启用拆分，可在 `main.py` 中取消以下注释：

```python
separate = Separate()
separate.generate(quotation_filename)
```

## 注意事项

- `openpyxl` 不会计算公式，只能读取 Excel 已保存的缓存结果。
- 如果未安装 Excel，又希望拆分后的文件中公式位置直接变成数值，结果可能不完整，取决于源文件是否已经被 Excel 计算并保存过。
- 当前项目中存在一些历史模块和注释代码，README 以当前 `main.py` 的实际行为为准。
