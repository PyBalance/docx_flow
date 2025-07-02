# -*- coding: utf-8 -*-
"""
Docx Enhanced Toolkit - 流畅选择器模块

包含流畅选择器类，支持链式调用。
"""

from typing import List, Any, TYPE_CHECKING

from .conditions import Condition, FunctionCondition

if TYPE_CHECKING:
    from .editor import DocxEditor


class FluentSelector:
    """流畅选择器，支持链式调用。"""
    def __init__(self, elements: List[Any], editor: 'DocxEditor'):
        self._elements = elements
        self._editor = editor

    def get_by_index(self, index: int) -> 'FluentSelector':
        # 把负索引转换为正索引
        if index < 0:
            index += len(self._elements)
        if 0 <= index < len(self._elements):
            return FluentSelector([self._elements[index]], self._editor)
        return FluentSelector([], self._editor)

    def where(self, condition: Condition) -> 'FluentSelector':
        """根据 Condition 对象筛选元素。"""
        filtered = [elem for elem in self._elements if condition.check(elem)]
        return FluentSelector(filtered, self._editor)

    def in_section(self, section_index: int) -> 'FluentSelector':
        """按节索引筛选元素的便捷方法。"""
        def check_func(element):
            return self._editor.get_element_section_index(element) == section_index
        return self.where(FunctionCondition(check_func))
    
    def from_section(self, start_section_index: int) -> 'FluentSelector':
        """从指定节开始筛选元素的便捷方法。"""
        def check_func(element):
            idx = self._editor.get_element_section_index(element)
            return idx is not None and idx >= start_section_index
        return self.where(FunctionCondition(check_func))

    def apply(self, action) -> 'FluentSelector':
        """将一个 Action 应用于所有当前选中的元素。"""
        if not self._elements:
            print("没有选中任何元素，无法执行操作。")
            
        for element in self._elements:
            action.execute(element)
        return self

    @property
    def count(self) -> int:
        """返回当前选中元素的数量。"""
        return len(self._elements)

    def get(self) -> List[Any]:
        """获取所有当前选中的元素。"""
        return self._elements