"""
后端代码补丁：
1. 修复 _clean_doubao_address (去掉破坏地址的正则) — 已手工完成
2. 修复 call_doubao_for_address (加入大湾区约束) — 已手工完成
3. 修改 search_heritage_address 中百度搜索关键词部分 (加大湾区)
4. 修改 fallback_search (加大湾区)
5. 修改 /api/export 输出列 (项目名称 | 确认地址 | 经度 | 纬度 | AI 置信度 | 备注)
"""
import re
import os

APP_PY = "/workspace/app.py"
with open(APP_PY, "r", encoding="utf-8") as f:
    src = f.read()

# === Patch 3.1: 在百度网页搜索前（如果还有 provinces 分支的话）替换为固定大湾区关键词 ===
# 查找 "第二优先：百度网页搜索" / "回退 2：百度网页搜索" 后的关键词构造
# 由于前面的 AI 段已手工替换，但百度关键词 construction 可能仍有 provinces 分支
# 这里做温和替换：把 search_query 构造的 whole block 换成大湾区版本

# 目标：替换 "如果 custom_keyword: search_query = custom_keyword  else ... provinces ..." 这样的构造
# 更简单：在百度搜索块内，找到包含 heritage_qs / provinces / 非遗 保护单位 地址 的部分
# 直接替换关键词构造段。我们用正则找"# 构造搜索关键词"或者"if custom_keyword:"开始直到 url= 之前的段

# 我们的目标：从百度 search 块（紧跟在 "回退 2" 或 "第二优先"注释之后）
# 的 keyword 构造段，替换为统一的大湾区版本

# 先尝试：如果还存在 "如果 provinces:" 的情况则替换
old_prov_branch = '''            if custom_keyword:
                search_query = custom_keyword
            else:
                # 使用非遗专用搜索词提高命中率
                extracted_kw = extract_search_keywords(keyword)
                heritage_qs = extracted_kw.get('heritage_queries', [])
                
                if provinces:
                    province_str = ' '.join([p.replace('省','').replace('市','').replace('自治区','')
                                              .replace('壮族','').replace('回族','').replace('维吾尔','')
                                              .replace('特别行政区','') for p in provinces[:3]])
                    best_query = ''
                    for q in heritage_qs[:5]:
                        query_with_province = f"{q} {province_str}"
                        best_query = query_with_province
                        break
                    if not best_query:
                        best_query = f"{keyword} 非遗 保护单位 地址"
                    search_query = best_query
                else:
                    search_query = heritage_qs[0] if heritage_qs else f"{keyword} 非遗 保护单位 地址"'''

new_gba_keywords = '''            if custom_keyword:
                search_query = _wrap_gba(custom_keyword)
            else:
                extracted_kw = extract_search_keywords(keyword)
                heritage_qs = extracted_kw.get('heritage_queries', [])
                if heritage_qs:
                    search_query = f"粤港澳大湾区 {heritage_qs[0]}"
                else:
                    search_query = f"粤港澳大湾区 {keyword} 非遗 保护单位 地址"'''

if old_prov_branch in src:
    print("[Patch 3.1] 替换百度搜索关键词构造...")
    src = src.replace(old_prov_branch, new_gba_keywords)
else:
    # 可能已经被替换过，跳过
    print("[Patch 3.1] 没找到旧 provinces 分支（可能已被替换），跳过")

# === Patch 3.2: 提取地址后，增加大湾区城市过滤 ===
# 找百度搜索块的结果循环中是否有 "address = " 的逻辑，末尾添加 "不包含大湾区城市 → 丢弃"
# 我们的替换目标：result_item 构造前，如果已有地址，检查是否在大湾区
# 这里采用轻量级策略：在 "result_item = {" 之前添加过滤（仅限百度 search 块）
# 由于前面编辑已经包含在结果中，我们这里跳过。

# === Patch 4: fallback_search 同样添加大湾区约束 ===
# 找 fallback_search 中 search_query 构造，并替换为带大湾区前缀
# 这部分较为简单：在 fallback_search 函数内，将任意最终搜索 query 用 _wrap_gba 包装
# 实现：在 fallback_search 函数定义后，找到它的搜索 query 并在前面插入大湾区前缀

# 用更简单方式：在 fallback_search 函数内替换关键词构造
# 先找到 fallback_search 函数中的关键词构造（可能还保留着省份分支）
old_fb_kw = '''        if custom_keyword:
            fb_search_query = custom_keyword
        else:
            extracted_kw = extract_search_keywords(keyword)
            heritage_qs = extracted_kw.get('heritage_queries', [])
            if heritage_qs:
                fb_search_query = heritage_qs[0]
            else:
                fb_search_query = f"{keyword} 非遗 保护单位 地址"'''

new_fb_kw = '''        if custom_keyword:
            fb_search_query = _wrap_gba(custom_keyword)
        else:
            extracted_kw = extract_search_keywords(keyword)
            heritage_qs = extracted_kw.get('heritage_queries', [])
            if heritage_qs:
                fb_search_query = f"粤港澳大湾区 {heritage_qs[0]}"
            else:
                fb_search_query = f"粤港澳大湾区 {keyword} 非遗 保护单位 地址"'''

if old_fb_kw in src:
    print("[Patch 4] 替换 fallback_search 关键词构造...")
    src = src.replace(old_fb_kw, new_fb_kw)
else:
    # fallback_search 可能使用别的变量名，让我们搜索 "def fallback_search" 后看内容
    m = re.search(r'def fallback_search.*?(?=\ndef |\Z)', src, re.DOTALL)
    if m:
        body = m.group(0)
        # 如果 body 中没包含 "粤港澳"，打印一下前 200 字方便调试
        if "粤港澳" not in body:
            print(f"[Patch 4] fallback_search 中没找到大湾区约束。函数头：{body[:200]}")
        else:
            print("[Patch 4] fallback_search 中已有大湾区约束，跳过")
    else:
        print("[Patch 4] 未找到 fallback_search 函数，跳过")

# === Patch 5: 修改 /api/export 输出列 ===
# 目标：将
#   row['确认地址'] = item.get('confirmed_address', '')
#   row['经度'] = item.get('confirmed_longitude', '')
#   row['纬度'] = item.get('confirmed_latitude', '')
#   row['处理状态'] = item.get('status', '')
# 替换为：
#   row['项目名称'] = item.get('name', item.get('project_name', ''))
#   row['确认地址'] = item.get('confirmed_address', '')
#   row['经度'] = item.get('confirmed_longitude', '')
#   row['纬度'] = item.get('confirmed_latitude', '')
#   row['AI 置信度'] = item.get('address_confidence', item.get('confidence', ''))
#   row['备注'] = item.get('note', '')

old_export = """            row['确认地址'] = item.get('confirmed_address', '')
            row['经度'] = item.get('confirmed_longitude', '')
            row['纬度'] = item.get('confirmed_latitude', '')
            row['处理状态'] = item.get('status', '')"""

new_export = """            row['项目名称'] = item.get('name', item.get('project_name', ''))
            row['确认地址'] = item.get('confirmed_address', '')
            row['经度'] = item.get('confirmed_longitude', '')
            row['纬度'] = item.get('confirmed_latitude', '')
            row['AI 置信度'] = item.get('address_confidence', item.get('confidence', ''))
            row['备注'] = item.get('note', '')"""

if old_export in src:
    print("[Patch 5] 替换导出 Excel 列字段...")
    src = src.replace(old_export, new_export)
else:
    print(f"[Patch 5] 没找到旧导出字段。当前 export 附近内容：")
    # 打印 export 附近内容
    idx = src.find("def export_excel")
    if idx >= 0:
        print(src[idx:idx + 500])

# === Patch 6: search_with_ai 接口中的豆包搜索，增加大湾区提示 ===
# 在 search_with_ai 的 prompt 中加入粤港澳大湾区约束
# 这里我们的 call_doubao_for_address 已经在之前的手工编辑中加过了，无需再改

# 保存
with open(APP_PY, "w", encoding="utf-8") as f:
    f.write(src)
print("Done.")
