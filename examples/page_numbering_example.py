#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
页码功能示例

演示如何使用 docx_flow 为 Word 文档添加和管理页码。
"""

from docx_flow import DocxEditor
from docx_flow.actions import AddPageNumberAction, ClearPageNumberAction


def demonstrate_page_numbering():
    """演示页码功能的各种用法"""
    print("🌟 Docx Flow 页码功能演示 🌟")
    print("=" * 50)
    
    # 假设我们有一个包含多个节的文档
    # 在实际使用中，请替换为您的文档路径
    input_file = "multi_section_document.docx"
    output_file = "page_numbered_output.docx"
    
    try:
        editor = DocxEditor(input_file)
        print(f"📖 已加载文档: {input_file}")
        print(f"📄 文档包含 {editor.select_sections().count} 个节")
        
        # 演示1: 清除所有现有页码
        print("\n--- 演示1: 清除所有页码 ---")
        editor.select_sections().apply(ClearPageNumberAction())
        print("✅ 已清除所有节的页码")
        
        # 演示2: 为所有节添加默认页码
        print("\n--- 演示2: 添加默认页码 ---")
        editor.select_sections().apply(AddPageNumberAction())
        print("✅ 已为所有节添加默认页码（微软雅黑9号，居中）")
        
        # 演示3: 从第二节开始连续编页码
        print("\n--- 演示3: 从第二节开始连续编页码 ---")
        # 先清除所有页码
        editor.select_sections().apply(ClearPageNumberAction())
        # 从第二节开始添加页码
        editor.select_sections().from_section(1)\
            .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
        print("✅ 第一节无页码，从第二节开始编页码")
        
        # 演示4: 分组页码编号
        print("\n--- 演示4: 分组页码编号 ---")
        # 清除所有页码
        editor.select_sections().apply(ClearPageNumberAction())
        
        # 第1-2节为第一组（页码1-x）
        if editor.select_sections().count >= 1:
            editor.select_sections().get_by_index(0)\
                .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
            print("✅ 第1节：重新开始编号，从1开始")
            
        if editor.select_sections().count >= 2:
            editor.select_sections().get_by_index(1)\
                .apply(AddPageNumberAction(restart_numbering=False))
            print("✅ 第2节：继续编号")
        
        # 第3节开始为第二组（重新从1开始）
        if editor.select_sections().count >= 3:
            editor.select_sections().from_section(2)\
                .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
            print("✅ 第3节及以后：重新开始编号，从1开始")
        
        # 演示5: 自定义页码格式
        print("\n--- 演示5: 自定义页码格式 ---")
        # 为第一节设置特殊格式的页码
        if editor.select_sections().count >= 1:
            editor.select_sections().get_by_index(0)\
                .apply(AddPageNumberAction(
                    start_number=1,
                    restart_numbering=True,
                    font_name='Arial',
                    font_size=10,
                    alignment='right'
                ))
            print("✅ 第1节：Arial字体，10号，右对齐")
        
        # 演示6: 为特定节范围添加页码
        print("\n--- 演示6: 为中间的节添加页码 ---")
        # 清除所有页码
        editor.select_sections().apply(ClearPageNumberAction())
        
        # 只为第2-3节添加页码
        section_count = editor.select_sections().count
        if section_count >= 3:
            # 第2节开始编号
            editor.select_sections().get_by_index(1)\
                .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
            # 第3节继续编号
            editor.select_sections().get_by_index(2)\
                .apply(AddPageNumberAction(restart_numbering=False))
            print("✅ 只有第2-3节有页码，其他节无页码")
        
        # 保存文档
        editor.save(output_file)
        print(f"\n💾 文档已保存至: {output_file}")
        print("🎉 页码功能演示完成！")
        
    except FileNotFoundError:
        print(f"❌ 找不到输入文件: {input_file}")
        print("请确保文件存在，或修改 input_file 变量指向正确的文档路径")
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")


def create_sample_scenarios():
    """创建一些常见的页码使用场景示例"""
    print("\n📚 常见页码使用场景")
    print("=" * 30)
    
    scenarios = [
        {
            "name": "学术论文",
            "description": "封面无页码，目录用罗马数字，正文从1开始",
            "code": """
# 清除所有页码
editor.select_sections().apply(ClearPageNumberAction())

# 封面（第1节）：无页码
# 目录（假设第2节）：暂不支持罗马数字，跳过或使用数字
# 正文（第3节开始）：从1开始
editor.select_sections().from_section(2)\\
    .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
"""
        },
        {
            "name": "技术手册",
            "description": "每个章节重新编页码",
            "code": """
# 为每个节都重新开始页码编号
for i in range(editor.select_sections().count):
    editor.select_sections().get_by_index(i)\\
        .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
"""
        },
        {
            "name": "合同文件",
            "description": "正文连续编页码，附件重新编号",
            "code": """
# 假设前3节是正文，后续节是附件
# 正文连续编号
editor.select_sections().get_by_index(0)\\
    .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
for i in range(1, 3):
    editor.select_sections().get_by_index(i)\\
        .apply(AddPageNumberAction(restart_numbering=False))

# 附件重新编号
editor.select_sections().from_section(3)\\
    .apply(AddPageNumberAction(start_number=1, restart_numbering=True))
"""
        }
    ]
    
    for scenario in scenarios:
        print(f"\n📋 {scenario['name']}")
        print(f"   {scenario['description']}")
        print(f"   代码示例:{scenario['code']}")


if __name__ == "__main__":
    demonstrate_page_numbering()
    create_sample_scenarios()