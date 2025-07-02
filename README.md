# docx-flow

一个基于 python-docx 的流式API库，用于对 Word 文档进行条件化修改和批量处理。

当前版本：v0.1.0

## 特性

- **流畅的链式调用接口**：使用直观的方法链来选择和修改文档元素
- **条件化选择**：基于正则表达式、表格列数或自定义函数来筛选元素
- **节级别操作**：支持按文档节来限制操作范围
- **丰富的操作类型**：支持段落对齐、表格格式化、文本替换、制表位设置等
- **类型安全**：完整的类型注解支持
- **全面测试**：确保代码质量和稳定性

## 安装

### 使用 uv (推荐)
```bash
uv add git+https://github.com/PyBalance/docx-flow.git
```

### 使用 pip
```bash
pip install git+https://github.com/PyBalance/docx-flow.git
```

### 使用 pipenv
```bash
pipenv install git+https://github.com/PyBalance/docx-flow.git#egg=docx-flow
```

### 更新到最新版本

```bash
# 使用 uv 更新
uv add --upgrade git+https://github.com/PyBalance/docx-flow.git

# 使用 pip 更新
pip install --upgrade git+https://github.com/PyBalance/docx-flow.git

# 使用 pipenv 更新
pipenv update docx-flow
```

## 快速开始

### 基本用法

```python
from docx_flow import DocxEditor, RegexCondition, AlignParagraphAction

# 打开文档
editor = DocxEditor('document.docx')

# 将包含"重要"的段落右对齐
editor.select_paragraphs()\
      .where(RegexCondition(r'重要'))\
      .apply(AlignParagraphAction('right'))

# 保存修改
editor.save('modified_document.docx')
```

### 复杂操作示例

```python
from docx_flow import (
    DocxEditor, RegexCondition, TableColumnCondition, TableTextCondition,
    AlignParagraphAction, SetTableWidthAction, ReplaceTextAction,
    SetTabStopAction, RemoveTableBordersAction, SetTableColumnWidthAction,
    AutoFitTableAction, SetFontSizeAction, ClearAndSetTabStopAction
)
from docx.shared import Inches

# 打开文档
editor = DocxEditor('complex_document.docx')

# 1. 第一节中的表格去掉边框，设置宽度
editor.select_tables()\
      .in_section(0)\
      .apply(RemoveTableBordersAction())\
      .apply(SetTableWidthAction(Inches(6.0)))

# 2. 第二节中包含"重要"的段落右对齐
editor.select_paragraphs()\
      .in_section(1)\
      .where(RegexCondition(r'重要'))\
      .apply(AlignParagraphAction('right'))

# 3. 第二节所有段落添加制表位
editor.select_paragraphs()\
      .in_section(1)\
      .apply(SetTabStopAction(2.0))

# 4. 从第三节开始替换文本
editor.select_paragraphs()\
      .from_section(2)\
      .apply(ReplaceTextAction('旧文本', '新文本'))

# 5. 4列表格设置列宽
col_widths = [Inches(1.5), Inches(2.0), Inches(1.5), Inches(2.0)]
editor.select_tables()\
      .where(TableColumnCondition(4))\
      .apply(SetTableColumnWidthAction(col_widths))

# 保存结果
editor.save('processed_document.docx')
```

## 核心概念

### 选择器 (Selectors)

选择器用于从文档中选择特定的元素：

- `select_paragraphs()` - 选择所有段落
- `select_tables()` - 选择所有表格
- `select_sections()` - 选择所有节

### 条件 (Conditions)

条件用于过滤选中的元素：

- `RegexCondition(pattern)` - 基于正则表达式匹配段落文本
- `TableColumnCondition(count)` - 基于列数匹配表格
- `TableTextCondition(text)` - 基于文本内容匹配表格
- `FunctionCondition(func)` - 基于自定义函数匹配元素

### 操作 (Actions)

操作用于修改选中的元素：

- `AlignParagraphAction(alignment)` - 设置段落对齐方式
- `SetTabStopAction(position_in_cm)` - 设置制表位（厘米）
- `ClearAndSetTabStopAction(position_in_cm)` - 清除并重设制表位（厘米）
- `ReplaceTextAction(old, new)` - 替换文本
- `SetFontSizeAction(size)` - 设置字体大小
- `SetTableWidthAction(width)` - 设置表格宽度
- `RemoveTableBordersAction()` - 移除表格边框
- `SetTableColumnWidthAction(widths)` - 设置表格列宽
- `AutoFitTableAction(mode)` - 自动调整表格大小

### 节级别过滤

- `in_section(index)` - 只处理指定节中的元素
- `from_section(index)` - 处理从指定节开始的所有元素

## API 参考

### DocxEditor

主要的编辑器类，用于加载和保存文档。

```python
class DocxEditor:
    def __init__(self, docx_path: str)
    def select_paragraphs(self) -> FluentSelector
    def select_tables(self) -> FluentSelector
    def select_sections(self) -> FluentSelector
    def save(self, output_path: str) -> None
```

### FluentSelector

流畅接口的核心类，支持链式调用。

```python
class FluentSelector:
    def where(self, condition: Condition) -> 'FluentSelector'
    def in_section(self, section_index: int) -> 'FluentSelector'
    def from_section(self, section_index: int) -> 'FluentSelector'
    def apply(self, action: Action) -> 'FluentSelector'
    def get(self) -> List[Any]
    def get_by_index(self, index: int) -> 'FluentSelector'
    @property
    def count(self) -> int
```

### 条件类

#### RegexCondition

```python
class RegexCondition(Condition):
    def __init__(self, pattern: str)
```

匹配段落文本的正则表达式条件。

#### TableColumnCondition

```python
class TableColumnCondition(Condition):
    def __init__(self, column_count: int)
```

匹配指定列数的表格条件。

#### TableTextCondition

```python
class TableTextCondition(Condition):
    def __init__(self, text: str)
```

匹配包含指定文本的表格条件。

#### FunctionCondition

```python
class FunctionCondition(Condition):
    def __init__(self, func: Callable[[Any], bool])
```

基于自定义函数的条件。

### 操作类

#### AlignParagraphAction

```python
class AlignParagraphAction(Action):
    def __init__(self, alignment: str)
```

设置段落对齐方式。支持的对齐方式：
- `'left'` - 左对齐
- `'center'` - 居中对齐
- `'right'` - 右对齐
- `'justify'` - 两端对齐

#### SetTabStopAction

```python
class SetTabStopAction(Action):
    def __init__(self, position_in_cm: float)
```

设置段落的制表位位置（以厘米为单位）。

#### SetFontSizeAction

```python
class SetFontSizeAction(Action):
    def __init__(self, size: Union[int, str])
```

设置字体大小。size可以是绝对值（如18）或相对值（如'+4'）。

#### ClearAndSetTabStopAction

```python
class ClearAndSetTabStopAction(Action):
    def __init__(self, position_in_cm: float)
```

清除现有制表位并设置新的制表位位置（以厘米为单位）。

#### AutoFitTableAction

```python
class AutoFitTableAction(Action):
    def __init__(self, autofit_mode: str = 'contents', first_col_ratio: float = None)
```

自动调整表格大小。支持'contents'（内容自适应）、'window'（窗口自适应）、'fixed'（固定宽度）模式。

#### ReplaceTextAction

```python
class ReplaceTextAction(Action):
    def __init__(self, old_text: str, new_text: str)
```

替换段落或表格单元格中的文本。

#### SetTableWidthAction

```python
class SetTableWidthAction(Action):
    def __init__(self, width: Any)
```

设置表格宽度。width 可以是 `Inches(6.0)`, `Cm(15.0)` 等。

#### RemoveTableBordersAction

```python
class RemoveTableBordersAction(Action):
```

移除表格的所有边框。

#### SetTableColumnWidthAction

```python
class SetTableColumnWidthAction(Action):
    def __init__(self, column_widths: List[Any])
```

设置表格各列的宽度。

## 使用场景

### 1. 文档格式标准化

```python
# 统一所有重要段落的格式
editor.select_paragraphs()\
      .where(RegexCondition(r'重要|关键|核心'))\
      .apply(AlignParagraphAction('center'))

# 统一所有表格的宽度
editor.select_tables()\
      .apply(SetTableWidthAction(Inches(6.0)))
```

### 2. 批量文本替换

```python
# 替换所有出现的公司名称
editor.select_paragraphs()\
      .apply(ReplaceTextAction('旧公司名', '新公司名'))

# 也替换表格中的内容
editor.select_tables()\
      .apply(ReplaceTextAction('旧公司名', '新公司名'))
```

### 3. 按节处理文档

```python
# 只处理第一节的内容
editor.select_paragraphs()\
      .in_section(0)\
      .apply(SetTabStopAction(2.0))

# 从第三节开始的所有内容
editor.select_paragraphs()\
      .from_section(2)\
      .apply(ReplaceTextAction('草稿', '正式版'))
```

### 4. 条件化表格处理

```python
# 只处理4列的表格
editor.select_tables()\
      .where(TableColumnCondition(4))\
      .apply(SetTableColumnWidthAction([
          Inches(1.0), Inches(2.0), Inches(1.5), Inches(1.5)
      ]))

# 移除所有表格的边框
editor.select_tables()\
      .apply(RemoveTableBordersAction())
```

## 注意事项

1. **文档备份**：在处理重要文档前，请务必备份原文件。

2. **测试环境**：建议先在测试文档上验证操作效果。

3. **性能考虑**：对于大型文档，某些操作可能需要较长时间。

4. **兼容性**：本库基于 python-docx，支持 .docx 格式（不支持 .doc）。

5. **错误处理**：库会优雅地处理错误情况，不匹配的操作会被跳过。

## 开发和测试

本项目使用 TDD（测试驱动开发）方式开发，包含完整的测试套件：

```bash
# 运行所有测试
python -m pytest

# 运行特定测试
python -m pytest tests/test_conditions.py

# 查看测试覆盖率
python -m pytest --cov=docx_toolkit
```

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 更新日志

### v0.1.0
- 初始版本发布
- 支持基本的段落和表格操作  
- 实现流畅的链式调用接口
- 完整的功能演示

