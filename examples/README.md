# Docx Flow 示例

这个目录包含了 Docx Flow 的各种使用示例，帮助您快速上手和了解库的功能。

## 📁 示例文件

### 1. `basic_usage.py` - 基础用法示例
演示最基本的功能：
- 创建示例文档
- 文本查找和替换
- 段落对齐和字体设置
- 表格自动调整
- 基本的条件筛选

**运行方式：**
```bash
cd examples
python basic_usage.py
```

**输出文件：**
- `sample_input.docx` - 原始示例文档
- `sample_output.docx` - 处理后的文档

### 2. `advanced_features.py` - 高级功能示例
演示更复杂的使用场景：
- 多节文档处理
- 自定义条件函数（如长段落识别）
- 按节筛选和处理
- 表格精确列宽设置
- 页面方向设置
- 复杂的条件组合

**运行方式：**
```bash
cd examples
python advanced_features.py
```

**输出文件：**
- `complex_input.docx` - 原始复杂文档
- `complex_output.docx` - 处理后的文档

## 🎯 学习路径建议

1. **先运行 `basic_usage.py`**
   - 了解基本的链式调用语法
   - 熟悉常用的条件和操作
   - 理解文档处理的基本流程

2. **然后运行 `advanced_features.py`**
   - 学习更复杂的筛选条件
   - 了解多节文档的处理方式
   - 掌握自定义函数条件的使用

3. **查看 `../demo.py`**
   - 这是最完整的演示文件
   - 包含了所有功能的综合展示
   - 可以作为功能参考手册

## 💡 使用技巧

### 条件组合
```python
# 可以链式使用多个条件
editor.select_paragraphs() \
    .where(RegexCondition(r'重要')) \
    .where(FunctionCondition(lambda p: len(p.text) > 20)) \
    .apply(action)
```

### 按节处理
```python
# 只处理特定节
editor.select_paragraphs().in_section(0)  # 第一节

# 从某节开始处理
editor.select_paragraphs().from_section(1)  # 第二节及以后
```

### 操作链
```python
# 可以对同一选择器应用多个操作
editor.select_paragraphs() \
    .where(condition) \
    .apply(ReplaceTextAction('旧', '新')) \
    .apply(AlignParagraphAction('center')) \
    .apply(SetFontSizeAction(14))
```

## 🔧 自定义扩展

您可以创建自己的条件和操作类：

### 自定义条件
```python
from docx_flow.conditions import Condition

class MyCustomCondition(Condition):
    def check(self, element):
        # 自定义筛选逻辑
        return True  # 或 False
```

### 自定义操作
```python
from docx_flow.actions import Action

class MyCustomAction(Action):
    def execute(self, element):
        # 自定义操作逻辑
        pass
```

## 📖 更多资源

- [主项目 README](../README.md) - 完整的 API 文档
- [demo.py](../demo.py) - 综合功能演示
- [源码](../docx_flow/) - 了解实现细节

如果您有问题或建议，欢迎提交 Issue！