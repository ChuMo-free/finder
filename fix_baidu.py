#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精准修复 app.py 中 search_heritage_address 的百度搜索部分
"""

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 找到需要修复的行范围
# 百度搜索部分从 "# 第二优先" 开始，到 "url =" 之前
# 需要把关键词构造逻辑替换成使用 keyword_variations

# 策略：找到  try:  后面的关键词构造部分，替换它
# 具体：找到 "search_query = f\"{keyword}" 这种行，替换逻辑

new_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    
    # 检测：是否进入了百度搜索的 try 块内部，且正在构造 search_query
    # 旧逻辑：
    #     if custom_keyword:
    #         search_query = custom_keyword
    #     else:
    #         if provinces:
    #             ...
    #         else:
    #             search_query = f"{keyword} 非遗 地址 传承基地 所在地"
    #
    # 新逻辑：
    #     search_query = keyword_variations[0]
    #     if provinces and not any(...):
    #         search_query = f"{search_query} {province_str}"
    
    if '# 第二优先：百度网页搜索' in line:
        # 找到这个注释，然后找到后面的 try: 块
        # 把 try: 块里面的关键词构造替换掉
        # 更好的方法：直接把从 try: 到 url = ... 之间的内容替换掉
        
        # 先找到 try: 的位置
        j = i
        while j < len(lines) and '    try:' not in lines[j]:
            new_lines.append(lines[j])
            j += 1
        
        # 现在 j 指向   try:
        new_lines.append(lines[j])  # 保留 try:
        j += 1
        
        # 跳过旧的关键词构造逻辑，直到找到   url = f"https://www.baidu.com/...
        while j < len(lines) and 'url = f"https://www.baidu.com/' not in lines[j] and 'url = f"https://www.baidu.com/' not in lines[j]:
            j += 1
        
        # 现在 j 指向 url = ... 这一行
        # 在 try: 和 url = ... 之间，插入新的关键词构造逻辑
        new_lines.append('        # 用第一个关键词变体搜索\n')
        new_lines.append('        search_query = keyword_variations[0]\n')
        new_lines.append('        \n')
        new_lines.append('        # 如果有省份限定，把省份加入关键词\n')
        new_lines.append('        if provinces and not any(p in search_query for p in provinces):\n')
        new_lines.append('            province_str = \" \".join([p.replace(\\'省\\',\').replace(\\'市\\',\').replace(\\'自治区\\',\'\')\n')
        new_lines.append('                                              .replace(\\'壮族\\',\').replace(\\'回族\\',\').replace(\\'维吾尔\\',\'\')\n')
        new_lines.append('                                              .replace(\\'特别行政区\\',\'\') for p in provinces[:3]])\n')
        new_lines.append('            search_query = f"{search_query} {province_str}"\n')
        new_lines.append('        \n')
        
        # 现在把 url = ... 这一行及之后的内容加回来
        while i < j:
            i += 1
        continue  # 跳到 url = ... 的处理
    
    new_lines.append(line)
    i += 1

# 写回文件
with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('[OK] 已修复百度搜索部分')
