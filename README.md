# Bidding Script

用于援外项目投标文件的自动化生成。

## 当前功能

- 从 `project*.docx` 读取项目数据（项目名称、招标编号、物资清单、服务需求等）
- 自动创建投标文件目录结构（`directory.py`）
- 自动生成投标目录 Excel（`content-<项目名>.xlsx`）
- 自动生成报价表 Excel（`quotation-<项目名>.xlsx`）
- 已包含基础容错：
  - 未找到 `project*.docx` 时抛出明确错误
  - 目录已存在时不会因 `FileExistsError` 中断

## 依赖环境

- Python 3.10+
- `python-docx`
- `openpyxl`

安装依赖：

```bash
pip install python-docx openpyxl
```

## 输入文件

将项目 Word 文件放在项目根目录，文件名需匹配：

- `project*.docx`

示例：

- `project-[Project Name].docx`

## 运行方式

```bash
python main.py
```

## 输出结果

运行后会在项目根目录生成：

- `投标文件-<项目名>/...`（目录结构）
- `content-<项目名>.xlsx`（投标文件目录）
- `quotation-<项目名>.xlsx`（报价表）

## 代码结构

- `main.py`：入口脚本
- `project.py`：读取并解析 Word 项目信息
- `directory.py`：创建投标目录结构
- `content.py`：生成目录与报价表
- `cover.py`：封面生成（当前入口未启用）

## 注意事项

- 当前正则会匹配目录下第一个 `project*.docx` 文件，建议同目录只放一个项目输入文件。
- 报价表中单价列默认留空，金额列通过公式自动计算。
