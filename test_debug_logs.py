#!/usr/bin/env python3
"""
测试脚本：验证空格处理的debug日志输出
用于调试空格丢失问题，提供完整的空格处理过程可观测性
"""

import logging
from utils.ooxml_rebuilder import OOXMLRebuilder

# Configure root logger to show debug messages
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def test_space_logic():
    """测试空格判断逻辑"""
    print("=" * 60)
    print("测试空格判断逻辑 (_should_add_space_after)")
    print("=" * 60)
    
    rebuilder = OOXMLRebuilder()
    
    # 测试用例
    test_cases = [
        ("hello", "world", True),   # 字母->字母：应该添加空格
        ("test", "123", True),      # 字母->数字：应该添加空格  
        ("123", "abc", True),       # 数字->字母：应该添加空格
        ("word", ".", False),       # 字母->标点：不应该添加空格
        (".", "word", False),       # 标点->字母：不应该添加空格
        ("hello ", "world", False), # 已有空格：不重复添加
        ("", "word", False),        # 空字符串：不添加空格
        ("word", "", False),        # 空字符串：不添加空格
    ]
    
    for i, (current, next_text, expected) in enumerate(test_cases, 1):
        print(f"\n--- 测试用例 {i} ---")
        result = rebuilder._should_add_space_after(current, next_text)
        status = "✓" if result == expected else "✗"
        print(f"结果: {result} (期望: {expected}) {status}")

def test_preprocessing_logic():
    """测试空格预处理逻辑"""
    print("\n" + "=" * 60)
    print("测试空格预处理逻辑 (_preprocess_segments_with_spaces)")
    print("=" * 60)
    
    rebuilder = OOXMLRebuilder()
    
    # 模拟segments数据
    test_segments = [
        {
            'sequence_id': 1,
            'translated_text': 'Hello',
            'original_text': 'Hello'
        },
        {
            'sequence_id': 2, 
            'translated_text': 'world',
            'original_text': 'world'
        },
        {
            'sequence_id': 3,
            'translated_text': 'test',
            'original_text': 'test'  
        },
        {
            'sequence_id': 4,
            'translated_text': '.',
            'original_text': '.'
        }
    ]
    
    print(f"\n输入segments: {len(test_segments)} 个")
    for seg in test_segments:
        print(f"  seq_{seg['sequence_id']}: {repr(seg['translated_text'])}")
    
    # 执行预处理
    processed_segments = rebuilder._preprocess_segments_with_spaces(test_segments)
    
    print(f"\n输出segments: {len(processed_segments)} 个")
    for seg in processed_segments:
        print(f"  seq_{seg['sequence_id']}: {repr(seg['translated_text'])}")

def test_xml_replacement_simulation():
    """模拟XML替换过程的日志输出"""
    print("\n" + "=" * 60)
    print("模拟XML替换过程日志")  
    print("=" * 60)
    
    rebuilder = OOXMLRebuilder()
    
    # 模拟已经预处理过的segments（包含添加的空格）
    mock_segments = [
        {
            'sequence_id': 1,
            'translated_text': 'Hello ',  # 注意末尾的空格
            'original_text': 'Hello',
            'xml_location': {
                'element_xpath': '//w:t[1]',
                'namespace_map': {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            }
        },
        {
            'sequence_id': 2,
            'translated_text': 'world',
            'original_text': 'world', 
            'xml_location': {
                'element_xpath': '//w:t[2]',
                'namespace_map': {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            }
        }
    ]
    
    print(f"\n模拟处理 {len(mock_segments)} 个segments的XML替换过程:")
    for i, seg in enumerate(mock_segments, 1):
        seq_id = seg['sequence_id']
        translated_text = seg['translated_text']
        xpath = seg['xml_location']['element_xpath']
        
        print(f"\n--- 模拟处理segment {i} (seq_{seq_id}) ---")
        print(f"XPath: {xpath}")
        print(f"翻译文本: {repr(translated_text)}")
        print(f"包含空格: {translated_text.count(' ')} 个")
        print(f"末尾空格: {translated_text.endswith(' ')}")

if __name__ == "__main__":
    print("🔧 开始测试空格处理Debug日志")
    print("这个脚本将展示所有添加的debug日志输出，帮助调试空格丢失问题")
    
    try:
        test_space_logic()
        test_preprocessing_logic()
        test_xml_replacement_simulation()
        
        print("\n" + "=" * 60)
        print("✅ 测试完成！")
        print("=" * 60)
        print("\n📋 如何使用这些debug日志调试空格问题：")
        print("1. 在实际使用插件时，确保日志级别设置为DEBUG")
        print("2. 观察 [OOXMLRebuilder] 前缀的日志消息")
        print("3. 重点关注以下关键信息：")
        print("   - 'Space check' 消息：显示空格判断过程")
        print("   - 'SPACE ADDED' 消息：显示何时添加了空格") 
        print("   - 'Space analysis' 消息：显示文本中的空格统计")
        print("   - 'SPACE LOST' 警告：显示XML替换过程中的空格丢失")
        print("4. 使用repr()格式的输出来查看不可见字符（如空格）")
        
    except Exception as e:
        print(f"\n❌ 测试过程中出现错误: {str(e)}")
        import traceback
        traceback.print_exc()