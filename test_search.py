#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试搜索功能
"""

import sys
import os

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import search_heritage_address, multi_level_search, smart_extract_address, extract_place_name

def test_basic_search():
    """测试基本搜索功能"""
    print("=" * 60)
    print("测试1：基本搜索（百度搜索）")
    print("=" * 60)
    
    keyword = "龙门米饼"
    print(f"搜索关键词: {keyword}")
    
    try:
        results = search_heritage_address(keyword)
        print(f"搜索结果数量: {len(results)}")
        
        for i, result in enumerate(results[:3]):
            print(f"\n结果 {i+1}:")
            print(f"  标题: {result.get('title', '')}")
            print(f"  摘要: {result.get('snippet', '')[:100]}...")
            print(f"  提取地址: {result.get('extracted_address', '')}")
            print(f"  智能地址: {result.get('smart_address', '')}")
            print(f"  置信度: {result.get('address_confidence', 0)}")
            print(f"  来源: {result.get('source_name', '')}")
            print(f"  来源类型: {result.get('source_type', '')}")
            
    except Exception as e:
        print(f"搜索出错: {e}")
        import traceback
        traceback.print_exc()

def test_multi_level_search():
    """测试多级搜索功能"""
    print("\n" + "=" * 60)
    print("测试2：多级搜索")
    print("=" * 60)
    
    keyword = "龙门米饼"
    print(f"搜索关键词: {keyword}")
    
    try:
        result = multi_level_search(keyword, min_confidence=0.6)
        print(f"搜索成功: {result['success']}")
        print(f"最佳地址: {result['best_address']}")
        print(f"最佳置信度: {result['best_confidence']}")
        print(f"使用级别: {result['level']}")
        print(f"结果状态: {result['result_status']}")
        print(f"提取地名: {result['place_name']}")
        print(f"尝试级别数: {result['total_levels_tried']}")
        
        print("\n搜索日志:")
        for log in result['search_log']:
            print(f"  第{log['level']}级: {log['description']}")
            print(f"    百度结果: {log['baidu_result_count']}条, 必应结果: {log['bing_result_count']}条")
            if log.get('best_address'):
                print(f"    最佳地址: {log['best_address']} (置信度: {log['best_confidence']:.2f})")
            if log.get('fail_reason'):
                print(f"    失败原因: {log['fail_reason']}")
                
    except Exception as e:
        print(f"多级搜索出错: {e}")
        import traceback
        traceback.print_exc()

def test_smart_extract():
    """测试智能地址提取"""
    print("\n" + "=" * 60)
    print("测试3：智能地址提取")
    print("=" * 60)
    
    test_cases = [
        ("龙门米饼制作技艺", "龙门米饼是广东省惠州市龙门县的传统美食，历史悠久。", ""),
        ("惠州客家山歌", "惠州客家山歌流传于广东省惠州市一带，是客家文化的重要组成部分。", ""),
        ("粤剧", "粤剧，又称广东大戏，是广东省的传统戏剧。", ""),
    ]
    
    for title, snippet, url in test_cases:
        print(f"\n标题: {title}")
        print(f"摘要: {snippet}")
        
        result = smart_extract_address(title, snippet, url)
        print(f"提取地址: {result['address']}")
        print(f"置信度: {result['confidence']}")
        print(f"来源: {result['source']}")
        print(f"组件: {result['components']}")

def test_place_name_extract():
    """测试地名提取"""
    print("\n" + "=" * 60)
    print("测试4：地名提取")
    print("=" * 60)
    
    test_keywords = [
        "龙门米饼制作技艺",
        "惠州市客家山歌",
        "东莞龙舟竞渡",
        "佛山剪纸",
        "传统技艺",
    ]
    
    for keyword in test_keywords:
        place = extract_place_name(keyword)
        print(f"'{keyword}' -> 提取地名: '{place}'")

if __name__ == '__main__':
    test_basic_search()
    test_multi_level_search()
    test_smart_extract()
    test_place_name_extract()
    
    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)
