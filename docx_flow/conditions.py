# -*- coding: utf-8 -*-
"""
Docx Enhanced Toolkit - 条件模块

包含所有的条件类，用于筛选文档元素。
"""

import re
from abc import ABC, abstractmethod
from typing import Any, Callable

from docx.table import Table
from docx.text.paragraph import Paragraph


class Condition(ABC):
    """'条件'接口 (抽象基类)"""
    @abstractmethod
    def check(self, element: Any) -> bool:
        """
        检查给定的元素是否满足本条件。
        :param element: 文档元素 (Paragraph, Table, Section 等)。
        :return: 如果满足条件，返回 True，否则返回 False。
        """
        pass


class RegexCondition(Condition):
    """正则表达式条件：检查段落文本是否匹配特定模式。"""
    def __init__(self, pattern: str):
        self.pattern = re.compile(pattern)

    def check(self, element: Any) -> bool:
        if isinstance(element, Paragraph):
            return bool(self.pattern.search(element.text))
        return False


class TableColumnCondition(Condition):
    """表格列数条件：检查表格是否具有指定的列数。"""
    def __init__(self, column_count: int):
        self.column_count = column_count

    def check(self, element: Any) -> bool:
        if isinstance(element, Table):
            return len(element.columns) == self.column_count
        return False


class TableTextCondition(Condition):
    """表格文本条件：检查表格是否包含特定文本。"""
    def __init__(self, text: str):
        self.text = text

    def check(self, element: Any) -> bool:
        if isinstance(element, Table):
            for row in element.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if self.text in paragraph.text:
                            return True
        return False


class FunctionCondition(Condition):
    """通用函数条件：使用一个自定义函数作为检查逻辑。"""
    def __init__(self, func: Callable[[Any], bool]):
        self.func = func

    def check(self, element: Any) -> bool:
        try:
            return self.func(element)
        except Exception:
            return False