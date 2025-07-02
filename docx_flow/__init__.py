# -*- coding: utf-8 -*-
"""
Docx Enhanced Toolkit - 增强版 Python Docx 处理库
"""

__version__ = "0.1.0"
__author__ = "Docx Toolkit Team"

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
    SetSectionOrientationAction
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
    "SetFontSizeAction"
]

