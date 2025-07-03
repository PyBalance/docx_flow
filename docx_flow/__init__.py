# -*- coding: utf-8 -*-
"""
Docx Flow - 强大而优雅的 Word 文档自动化处理库

这个库提供了一套流畅的 API，让 Word 文档的批量处理变得简单直观。
支持链式调用，条件筛选，以及丰富的文档操作功能。

主要特性:
- 🔗 流畅的链式调用接口
- 🎯 强大的条件筛选功能 
- 📊 全面的表格处理能力
- 📝 灵活的段落操作
- 📄 完整的页面设置功能
- 🧩 模块化的设计架构

基本用法:
    from docx_flow import DocxEditor
    from docx_flow.conditions import RegexCondition
    from docx_flow.actions import ReplaceTextAction
    
    editor = DocxEditor('document.docx')
    editor.select_paragraphs() \\
        .where(RegexCondition(r'重要')) \\
        .apply(ReplaceTextAction('重要', '⭐ 重要'))
    editor.save('output.docx')

更多示例请参考项目文档和 demo.py 演示文件。
"""

__version__ = "0.1.1"
__author__ = "Docx Flow Team"
__description__ = "强大而优雅的 Word 文档自动化处理库"

from .editor import DocxEditor
from .selector import FluentSelector
from .conditions import (
    Condition,
    RegexCondition,
    TableColumnCondition,
    TableTextCondition,
    FunctionCondition
)
from .actions import (
    Action,
    RemoveTableBordersAction,
    SetTableWidthAction,
    AutoFitTableAction,
    SetTableColumnWidthAction,
    AlignParagraphAction,
    SetTabStopAction,
    ClearAndSetTabStopAction,
    ReplaceTextAction,
    SetFontSizeAction,
    SetSectionOrientationAction,
    AddPageNumberAction,
    ClearPageNumberAction
)

__all__ = [
    "DocxEditor",
    "FluentSelector",
    "Condition",
    "Action",
    "RegexCondition",
    "TableColumnCondition",
    "FunctionCondition",
    "RemoveTableBordersAction",
    "SetTableWidthAction",
    "AlignParagraphAction",
    "SetTabStopAction",
    "ReplaceTextAction",
    "SetSectionOrientationAction",
    "SetTableColumnWidthAction",
    "ClearAndSetTabStopAction",
    "AutoFitTableAction",
    "TableTextCondition",
    "SetFontSizeAction",
    "AddPageNumberAction",
    "ClearPageNumberAction"
]

