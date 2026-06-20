"""
非遗地址查找工具 - Flask后端应用
功能：Excel上传 → 字段映射 → 网络搜索地址 → 人工审核 → 高德坐标转换 → 人工审核 → 导出Excel
"""

import os
import uuid
import threading
import time
import json
import re
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ==================== AI模型配置（豆包） ====================
# 豆包API配置 - API Key可从前端传入，也可在这里填写默认值
# 获取方式：https://www.volcengine.com/docs/82379
# 推理接入点ID在火山引擎控制台获取，格式如：ep-20250101xxxxxx-xxxxx
DOUBAO_API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
DOUBAO_MODEL = ""  # 替换为你的推理接入点ID，如：ep-20250620001700-4hgdm


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB限制

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ==================== 全局任务状态管理 ====================
# 任务状态字典，key为task_id
tasks = {}
tasks_lock = threading.Lock()

# 任务状态枚举
TASK_STATUS = {
    'PENDING': '待开始',
    'UPLOADED': '文件已上传',
    'MAPPED': '字段已映射',
    'SEARCHING': '正在搜索地址',
    'WAITING_ADDRESS_REVIEW': '等待地址审核',
    'GEOCODING': '正在转换坐标',
    'WAITING_COORD_REVIEW': '等待坐标审核',
    'COMPLETED': '已完成',
    'PAUSED': '已暂停',
    'INTERRUPTED': '已中断',
    'ERROR': '出错'
}

# 单条数据状态
ITEM_STATUS = {
    'PENDING': '待处理',
    'SEARCHING': '搜索中',
    'SEARCH_DONE': '搜索完成',
    'WAITING_ADDRESS_REVIEW': '等待地址审核',
    'ADDRESS_APPROVED': '地址已确认',
    'GEOCODING': '坐标转换中',
    'GEO_DONE': '坐标转换完成',
    'WAITING_COORD_REVIEW': '等待坐标审核',
    'COORD_APPROVED': '坐标已确认',
    'COMPLETED': '已完成',
    'ERROR': '出错',
    # 两遍模式新增状态
    'PASS1_AUTO_APPROVED': '第一遍自动确认',
    'PASS1_DONE': '第一遍完成',
    'PASS2_REVIEWING': '第二遍审核中',
    'PASS2_DONE': '第二遍完成',
}


# ==================== 工具函数 ====================
def generate_task_id():
    """生成唯一任务ID"""
    return str(uuid.uuid4())[:8]


def get_task(task_id):
    """获取任务状态"""
    with tasks_lock:
        return tasks.get(task_id)


def update_task(task_id, **kwargs):
    """更新任务状态"""
    with tasks_lock:
        if task_id in tasks:
            tasks[task_id].update(kwargs)
            tasks[task_id]['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def add_log(task_id, message, level='info'):
    """添加日志"""
    with tasks_lock:
        if task_id in tasks:
            log_entry = {
                'time': datetime.now().strftime('%H:%M:%S'),
                'level': level,
                'message': message
            }
            tasks[task_id]['logs'].append(log_entry)
            # 最多保留500条日志
            if len(tasks[task_id]['logs']) > 500:
                tasks[task_id]['logs'] = tasks[task_id]['logs'][-500:]


# ==================== 网络搜索功能 ====================
def extract_search_keywords(project_name):
    """
    从非遗项目名称中提取最佳搜索关键词（v3-非遗专用版）
    
    策略：针对非遗项目的特殊性，生成多组搜索词：
    - 保护单位查询词（最可能找到地址）
    - 传承人/传承基地查询词
    - 地理位置查询词
    - 短名（用于高德地理编码）
    
    返回: dict with keys: primary, short_name, bracket_name, heritage_queries, gaode_keywords, variations
    """
    original = project_name.strip()
    
    # 1. 提取括号里的具体名称
    bracket_content = ''
    match = re.search(r'[\(\（](.*?)[\)\）]', original)
    if match:
        bracket_content = match.group(1).strip()
    
    # 2. 去掉通用后缀，获取核心名
    generic_suffixes = [
        '制作技艺', '传统技艺', '传承基地', '简介', '介绍',
        '官网', '地址在哪里', '地址', '在哪里', '位置',
        '制作方法', '技艺', '文化', '传统', '工艺',
        '酿造技艺', '炮制技艺'
    ]
    
    def strip_suffixes(name):
        result = name
        for suf in generic_suffixes:
            if result.endswith(suf):
                result = result[:len(result) - len(suf)]
        return result.strip()
    
    # 3. 确定核心名称
    core_names = []
    if bracket_content and len(bracket_content) >= 2:
        core_names.append(strip_suffixes(bracket_content))
    core_full = strip_suffixes(original)
    if core_full not in core_names:
        core_names.append(core_full)
    if bracket_content and bracket_content not in core_names:
        core_names.insert(0, bracket_content)
    
    short_name = core_names[0] if core_names else original
    
    # 4. 构建非遗专用搜索词
    heritage_queries = []
    for core in core_names[:2]:
        heritage_queries.append(f"{core} 非遗 保护单位 地址")
        heritage_queries.append(f"{core} 非物质文化遗产 保护单位")
        heritage_queries.append(f"{core} 传承基地 地址")
        heritage_queries.append(f"{core} 传承人 工作单位")
        heritage_queries.append(f"{core} 生产地 在哪里")
        heritage_queries.append(f"{core} 制作 地址")
        heritage_queries.append(f"{core} 非遗 在哪里")
        heritage_queries.append(f"{core} 位于 哪里")
    
    # 5. 高德地图专用关键词
    gaode_keywords = []
    for core in core_names[:2]:
        gaode_keywords.append(core)
        gaode_keywords.append(f"{core} 非遗")
        gaode_keywords.append(f"{core} 传承基地")
        gaode_keywords.append(f"{core} 博物馆")
        gaode_keywords.append(f"{core} 展示馆")
    
    # 6. 构建通用变体
    variations = list(heritage_queries[:5])
    for v in [short_name, f"{short_name} 地址", original]:
        if v not in variations:
            variations.append(v)
    
    return {
        'primary': short_name,
        'short_name': short_name,
        'bracket_name': bracket_content,
        'core_names': core_names,
        'heritage_queries': heritage_queries,
        'gaode_keywords': gaode_keywords[:5],
        'variations': variations[:8]
    }



def search_amap_places(keyword, amap_key, provinces=None):
    """
    使用高德API搜索非遗项目地址（v3-智能版）
    
    改进策略：
    1. 先用核心名+"非遗/保护单位"搜索
    2. 再用纯文本搜索+严格相关性过滤
    3. 排除只匹配常见地名（如"九江"）的无关结果
    """
    results = []
    
    def _do_search(search_kw, types=None, city=''):
        try:
            url = "https://restapi.amap.com/v3/place/text"
            params = {
                'key': amap_key,
                'keywords': search_kw,
                'offset': 10,
                'page': 1,
                'extensions': 'all',
                'output': 'json'
            }
            if types:
                params['types'] = types
            if city:
                params['city'] = city
                params['citylimit'] = 'true'
            response = requests.get(url, params=params, timeout=10)
            data = response.json()
            if data.get('status') == '1' and 'pois' in data:
                return data['pois']
        except Exception as e:
            add_log_for_search(keyword, f"高德搜索错误[{search_kw}]: {str(e)}")
        return []
    
    def _is_poi_relevant(poi_name, poi_address, keyword_core):
        """判断POI是否与关键词真正相关（严格过滤无关地名）"""
        keyword_core = keyword_core.strip()
        if not keyword_core or len(keyword_core) < 2:
            return False, 0
        
        # 完全匹配
        if keyword_core in poi_name:
            return True, 1.0
        
        # 提取关键词中独特的字（排除常见地名用字）
        common_place_chars = {'市','县','区','镇','乡','村','街','路','城',
                             '省','江','山','河','湖','海','南','北','东','西'}
        unique_chars = [c for c in keyword_core 
                       if c not in common_place_chars and not c.isascii() and u'\u4e00' <= c <= u'\u9fff']
        
        if not unique_chars:
            unique_chars = list(keyword_core[:2])
        
        # POI名称中独特字符匹配度 >= 50%
        matched = sum(1 for c in unique_chars if c in poi_name)
        ratio = matched / len(unique_chars)
        if ratio >= 0.5:
            return True, round(ratio + 0.2, 2)
        
        # 地址中匹配
        addr_matched = sum(1 for c in unique_chars if c in poi_address)
        if addr_matched >= len(unique_chars) * 0.5 and poi_address:
            return True, round(addr_matched / len(unique_chars), 2)
        
        return False, 0
    
    def _format_poi_result(poi, keyword, relevance_score=0.5):
        name = poi.get('name', '')
        address = poi.get('address', '')
        location = poi.get('location', '')
        pname = poi.get('pname', '')
        cityname = poi.get('cityname', '')
        adname = poi.get('adname', '')
        
        full_address = f"{pname}{cityname}{adname}{address}" if address else f"{pname}{cityname}{adname}{name}"
        confidence = 0.3 + relevance_score * 0.4
        if address:
            confidence += 0.15
        if location:
            confidence += 0.15
        
        return {
            'index': 0,
            'title': f"📍 高德地图：{name}",
            'snippet': f"地址：{full_address}",
            'link': f"https://uri.amap.com/marker?position={location}&name={requests.utils.quote(name)}" if location else '',
            'source_url': 'https://lbs.amap.com/',
            'source_name': '高德地图API',
            'source_type': 'gaode',
            'extracted_address': full_address,
            'smart_address': full_address,
            'address_confidence': round(min(confidence, 1.0), 2),
            'address_components': {
                'province': pname, 'city': cityname, 'district': adname,
                'street': address,
                'lng': location.split(',')[0] if location else '',
                'lat': location.split(',')[1] if location else ''
            },
            'address_source': 'gaode_api',
            'source': '高德地图API',
            'gaode_location': location,
            'gaode_name': name,
            'relevance_score': relevance_score
        }
    
    # 城市参数
    city = ''
    if provinces and len(provinces) == 1:
        region = provinces[0].replace('省','').replace('市','').replace('自治区','').replace('壮族','').replace('回族','').replace('维吾尔','').replace('特别行政区','')
        city = region
    
    extracted = extract_search_keywords(keyword)
    core_keyword = extracted.get('short_name', keyword)
    seen_names = set()
    
    # 策略1: 核心+"非遗/保护单位"搜索
    for search_kw in [f"{core_keyword} 非遗", f"{core_keyword} 传承", f"{core_keyword} 保护"]:
        pois = _do_search(search_kw, city=city)
        for poi in pois:
            name = poi.get('name', '')
            if name in seen_names:
                continue
            is_rel, score = _is_poi_relevant(name, poi.get('address',''), core_keyword)
            if is_rel:
                seen_names.add(name)
                results.append(_format_poi_result(poi, keyword, score))
                if len(results) >= 5:
                    return results
    
    # 策略2: 纯文本搜索 + 严格过滤
    if len(results) < 3:
        pois = _do_search(core_keyword, city=city)
        for poi in pois:
            name = poi.get('name', '')
            if name in seen_names:
                continue
            is_rel, score = _is_poi_relevant(name, poi.get('address',''), core_keyword)
            if is_rel and score >= 0.5:
                seen_names.add(name)
                results.append(_format_poi_result(poi, keyword, score))
                if len(results) >= 5:
                    return results
    
    # 策略3: 其他关键词变体
    if len(results) < 2:
        for search_kw in extracted.get('gaode_keywords', []):
            if search_kw == core_keyword:
                continue
            pois = _do_search(search_kw, city=city)
            for poi in pois:
                name = poi.get('name', '')
                if name in seen_names:
                    continue
                is_rel, score = _is_poi_relevant(name, poi.get('address',''), core_keyword)
                if is_rel:
                    seen_names.add(name)
                    results.append(_format_poi_result(poi, keyword, score))
                    if len(results) >= 5:
                        return results
    
    return results



def search_heritage_address(keyword, custom_keyword=None, provinces=None, amap_key=None, doubao_api_key=None, doubao_model_id=None):
    """
    搜索非遗项目的真实地址
    优先使用高德地点搜索API，失败后再用网页搜索
    参数:
        keyword: 项目名称
        custom_keyword: 自定义搜索关键词（可选）
        provinces: 省份列表（可选），限定搜索省份以提高准确度
        amap_key: 高德API Key（可选），用于高德地点搜索
    """
    results = []
    
    # 第零优先：豆包AI搜索（如果配置了Key）
    if doubao_api_key:
        try:
            ai_result = call_doubao_for_address(keyword, doubao_api_key, amap_key, model_id=doubao_model_id)
        except Exception as e:
            ai_result = None
        if ai_result and ai_result.get('address'):
            results.append({
                'title': '豆包AI：' + str(ai_result.get('raw_response', '')),
                'snippet': 'AI直接回答，置信度：' + str(ai_result.get('confidence', '')),
                'address': ai_result['address'],
                'lng': ai_result.get('lng'),
                'lat': ai_result.get('lat'),
                'confidence': ai_result['confidence'],
                'source': ai_result['source'],
                'is_ai': True
            })
            # 高置信度直接返回，否则继续网页搜索
            if ai_result['confidence'] >= 0.9:
                return results
    
    
    # 提取多个搜索关键词变体
    if custom_keyword:
        keyword_variations = [custom_keyword]
    else:
        extracted = extract_search_keywords(keyword)
        keyword_variations = extracted['variations']
    
    # 第一优先：使用高德地点搜索API（如果用户已配置Key）
    if amap_key:
        for kw in keyword_variations:
            if results:
                break  # 已经找到结果，不再尝试其他关键词
            amap_results = search_amap_places(kw, amap_key, provinces)
            if amap_results:
                results.extend(amap_results)
                # 高德API结果足够好，直接返回（最多5条）
                if len(results) >= 3:
                    return results[:5]
    
    # 第二优先：百度网页搜索
    if len(results) < 3:
        search_query = keyword_variations[0]
        try:
            # 构造搜索关键词
            if custom_keyword:
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
                    search_query = heritage_qs[0] if heritage_qs else f"{keyword} 非遗 保护单位 地址"
            url = f"https://www.baidu.com/s?wd={requests.utils.quote(search_query)}"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9',
            }
        
            response = requests.get(url, headers=headers, timeout=10)
            response.encoding = 'utf-8'
        
            # CAPTCHA/封禁检测：如果百度返回验证码页面，跳过百度搜索
            html_lower = response.text.lower()
            if any(kw in html_lower for kw in ['验证码', 'captcha', 'unusual traffic', 'blocked', '请验证']):
                search_results = []  # 跳过百度，后续会尝试必应
            else:
                soup = BeautifulSoup(response.text, 'lxml')
            
                # 尝试多个选择器（百度HTML结构多次变化）
                search_results = []
                for selector in ['.c-container', '.result', '#content_left > div[tpl]', 'div[tpl="www_normal"]']:
                    search_results = soup.select(selector)
                    if len(search_results) >= 3:
                        break
            
                for idx, result in enumerate(search_results[:5]):
                    title_elem = result.select_one('h3 a, .t a')
                    snippet_elem = result.select_one('.c-abstract, .c-span9, .content-right_8Zs40')
                
                    title = title_elem.get_text(strip=True) if title_elem else ''
                    snippet = snippet_elem.get_text(strip=True) if snippet_elem else ''
                    link = title_elem.get('href', '') if title_elem else ''
                
                    # 获取真实URL
                    real_url = ''
                    if link and not link.startswith('http'):
                        # 百度跳转链接，尝试获取真实地址
                        try:
                            head_resp = requests.head(link, headers=headers, timeout=5, allow_redirects=True)
                            real_url = head_resp.url
                        except Exception:
                            real_url = link
                    elif link:
                        real_url = link
                
                    source_name = get_source_name(real_url) if real_url else '百度搜索'
                    source_type = get_source_type(source_name, real_url)
                
                    # v4: 先用非遗元数据提取（高精度）
                    heritage_result = extract_heritage_info(title + ' ' + snippet)
                    h_addr = heritage_result.get('address', '')
                    h_conf = heritage_result.get('confidence', 0)
                    
                    smart_result = smart_extract_address(title, snippet, real_url)
                    s_addr = smart_result['address']
                    s_conf = smart_result['confidence']
                    
                    # 选择更优的地址源：heritage > smart > fallback
                    if h_addr and h_conf >= s_conf:
                        address = h_addr
                        confidence = h_conf
                        addr_source = f'heritage:{heritage_result.get("source", "")}'
                        addr_components = dict(smart_result['components'])
                        addr_components['heritage_region'] = heritage_result.get('region_detail', '')
                        addr_components['organization'] = heritage_result.get('organization', '')
                    elif s_addr:
                        address = s_addr
                        confidence = s_conf
                        addr_source = smart_result['source']
                        addr_components = smart_result['components']
                    else:
                        # 最后尝试旧方法
                        old_address = extract_address_from_text(title + ' ' + snippet)
                        address = old_address if old_address else ''
                        confidence = 0.35 if address else 0
                        addr_source = 'legacy_fallback'
                        addr_components = {}
                    
                    # v4: 如果有地址且没有坐标，尝试地理编码
                    geo_location = None
                    if address and amap_key:
                        geo_result = geocode_text_address(address, amap_key)
                        if geo_result and geo_result.get('lng'):
                            geo_location = f"{geo_result['lng']},{geo_result['lat']}"
                            # 用地理编码返回的格式化地址（可能更标准）
                            if geo_result.get('formatted_address') and len(geo_result['formatted_address']) > len(address):
                                pass  # 保留原始提取地址，但记录格式化地址
                            confidence = max(confidence, geo_result.get('confidence', 0) * 0.9 + 0.1)
                            addr_components['geocoded'] = 'true'
                            addr_components['geo_province'] = geo_result.get('province', '')
                            addr_components['geo_city'] = geo_result.get('city', '')
                            addr_components['geo_district'] = geo_result.get('district', '')
                
                    result_item = {
                        'index': len(results) + 1,
                        'title': title,
                        'snippet': snippet,
                        'link': link,
                        'source_url': real_url,
                        'source_name': source_name,
                        'source_type': source_type,
                        'extracted_address': address,
                        'smart_address': address,
                        'address_confidence': round(confidence, 2),
                        'address_components': addr_components,
                        'address_source': addr_source,
                        'source': source_name
                    }
                    # 地理编码坐标
                    if geo_location:
                        result_item['gaode_location'] = geo_location
                        result_item['geo_formatted_address'] = geo_result.get('formatted_address', '') if geo_result else ''
                    
                    results.append(result_item)
    
        except Exception as e:
            # 百度搜索失败，静默跳过，后续尝试必应搜索
            pass
    # 如果百度搜索结果不足3条，尝试必应搜索
    if len(results) < 3:
        try:
            bing_results = fallback_search(keyword, custom_keyword=custom_keyword, provinces=provinces, amap_key=amap_key)
            # 只添加百度搜索没找到的结果
            existing_titles = {r.get('title', '') for r in results}
            for r in bing_results:
                if r.get('title', '') not in existing_titles:
                    r['index'] = len(results) + 1
                    results.append(r)
        except Exception:
            pass
    
    # 对搜索结果进行相关性过滤和排序
    if results:
        place_name = extract_place_name(keyword)
        results = filter_and_sort_results(results, keyword, place_name)
    
    return results[:5]


def fallback_search(keyword, custom_keyword=None, provinces=None, amap_key=None):
    """
    备用搜索方式：必应搜索 + 高德API（如有Key）
    参数:
        keyword: 项目名称
        custom_keyword: 自定义搜索关键词（可选）
        provinces: 省份列表（可选）
        amap_key: 高德API Key（可选）
    """
    results = []
    
    # 如果有高德Key，先尝试高德地点搜索
    if amap_key:
        gaode_results = search_amap_places(keyword, amap_key, provinces)
        if gaode_results:
            results.extend(gaode_results)
    
    # 必应搜索作为补充
    try:
        if custom_keyword:
            search_query = custom_keyword
        else:
            if provinces:
                province_str = ' '.join([p.replace('省','').replace('市','').replace('自治区','')
                                          .replace('壮族','').replace('回族','').replace('维吾尔','')
                                          .replace('特别行政区','') for p in provinces[:3]])
                fb_kw = extract_search_keywords(keyword)
                fb_qs = fb_kw.get('heritage_queries', [])
                search_query = f"{fb_qs[0]} {province_str}" if fb_qs else f"{keyword} {province_str} 非遗 保护单位 地址"
            else:
                fb_kw = extract_search_keywords(keyword)
                fb_qs = fb_kw.get('heritage_queries', [])
                search_query = fb_qs[0] if fb_qs else f"{keyword} 非遗 保护单位 地址"
        url = f"https://cn.bing.com/search?q={requests.utils.quote(search_query)}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'lxml')
        
        # 尝试多个必应选择器（必应HTML结构也会变化）
        items = []
        for selector in ['.b_algo', 'li.b_algo', '.b_title', '#b_results > li']:
            items = soup.select(selector)
            if len(items) >= 3:
                break
        
        for idx, item in enumerate(items[:5]):
            title_elem = item.select_one('h2 a, .b_title a')
            snippet_elem = item.select_one('.b_caption p, .b_algoSlug')
            
            title = title_elem.get_text(strip=True) if title_elem else ''
            snippet = snippet_elem.get_text(strip=True) if snippet_elem else ''
            link = title_elem.get('href', '') if title_elem else ''
            
            real_url = link
            
            source_name = get_source_name(real_url) if real_url else '必应搜索'
            source_type = get_source_type(source_name, real_url)
            
            # v4: 非遗元数据提取优先
            full_txt = title + (' ' + snippet if snippet else '')
            heritage_result = extract_heritage_info(full_txt)
            h_addr = heritage_result.get('address', '')
            h_conf = heritage_result.get('confidence', 0)
            
            smart_result = smart_extract_address(title, snippet, real_url)
            s_addr = smart_result['address']
            s_conf = smart_result['confidence']
            
            if h_addr and h_conf >= s_conf:
                address = h_addr
                confidence = h_conf
                addr_source = f'heritage:{heritage_result.get("source", "")}'
                addr_components = dict(smart_result['components'])
                addr_components['heritage_region'] = heritage_result.get('region_detail', '')
                addr_components['organization'] = heritage_result.get('organization', '')
            elif s_addr:
                address = s_addr
                confidence = s_conf
                addr_source = smart_result['source']
                addr_components = smart_result['components']
            else:
                old_address = extract_address_from_text(full_txt)
                address = old_address if old_address else ''
                confidence = 0.35 if address else 0
                addr_source = 'legacy_fallback'
                addr_components = {}
            
            # v4: 地理编码
            geo_location = None
            if address and amap_key:
                geo_result = geocode_text_address(address, amap_key)
                if geo_result and geo_result.get('lng'):
                    geo_location = f"{geo_result['lng']},{geo_result['lat']}"
                    confidence = max(confidence, geo_result.get('confidence', 0) * 0.9 + 0.1)
                    addr_components['geocoded'] = 'true'
                    addr_components['geo_province'] = geo_result.get('province', '')
                    addr_components['geo_city'] = geo_result.get('city', '')
                    addr_components['geo_district'] = geo_result.get('district', '')
            
            # 避免重复添加（已有高德结果时）
            if results and any(r.get('title', '') == title for r in results):
                continue
            
            result_item = {
                'index': len(results) + 1,
                'title': title,
                'snippet': snippet,
                'link': link,
                'source_url': real_url,
                'source_name': source_name,
                'source_type': source_type,
                'extracted_address': address,
                'smart_address': address,
                'address_confidence': round(confidence, 2),
                'address_components': addr_components,
                'address_source': addr_source,
                'source': source_name
            }
            if geo_location:
                result_item['gaode_location'] = geo_location
            
            results.append(result_item)
    except Exception:
        pass
    
    # 对搜索结果进行相关性过滤和排序
    if results:
        place_name = extract_place_name(keyword)
        results = filter_and_sort_results(results, keyword, place_name)
    
    return results[:5]





# ==================== 豆包AI搜索 ====================

def call_doubao_for_address(project_name, doubao_api_key, amap_key=None, model_id=None):
    """
    调用豆包AI API（Chat Completions接口，OpenAI兼容格式）直接获取非遗项目地址
    接口文档：https://www.volcengine.com/docs/82379
    参数:
        project_name: 非遗项目名称
        doubao_api_key: 豆包API Key（Bearer Token）
        amap_key: 高德API Key（可选，用于地理编码）
        model_id: 推理接入点ID，如 ep-20250620001700-4hgdm
    返回: {address, lng, lat, confidence, source, raw_response}
    """
    if not doubao_api_key:
        return None

    use_model = model_id if model_id else DOUBAO_MODEL

    # 构造提示词（让AI扮演非遗专家，直接回答地址）
    prompt = (
        "你是非遗数据专家。请回答：国家级或省级非遗项目"
        f"「{project_name}」\n"
        "的申报地区（或主要流传区域、保护单位所在地）是哪里？\n"
        "只回答标准地名（格式：省+市+县/区+镇/街道，例如：广东省佛山市南海区九江镇），"
        "不要加任何解释、标点符号或换行。如果不确定，回答「不确定」。"
    )

    try:
        headers = {
            "Authorization": f"Bearer {doubao_api_key}",
            "Content-Type": "application/json"
        }
        # 使用标准 OpenAI 兼容的 chat/completions 接口（火山引擎ARK稳定版）
        payload = {
            "model": use_model,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "stream": False,
            # 部分豆包模型支持联网搜索（需模型本身支持）
            "tools": [{"type": "web_search"}] if _model_supports_search(use_model) else None
        }
        # 移除None值的tools字段
        if payload["tools"] is None:
            del payload["tools"]

        print(f"[豆包AI] 请求URL: {DOUBAO_API_URL}, model: {use_model}")
        resp = requests.post(DOUBAO_API_URL, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        # 解析响应（标准 OpenAI chat/completions 格式）
        raw_text = ""
        print(f"[豆包AI] 收到响应，keys={list(data.keys())}")

        # 标准格式: {"choices": [{"message": {"content": "..."}}]}
        choices = data.get("choices", [])
        if choices and len(choices) > 0:
            msg = choices[0].get("message", {})
            raw_text = msg.get("content", "").strip()
        elif "output" in data:
            # Responses API 格式备用解析
            output = data["output"]
            if isinstance(output, list) and output:
                for item in output:
                    if isinstance(item, dict):
                        if item.get("type") == "message":
                            for c in item.get("content", []):
                                if isinstance(c, dict) and c.get("type") == "output_text":
                                    raw_text = c.get("text", "").strip()
                                    break

        if not raw_text:
            print(f"[豆包AI] 响应解析失败，原始响应前500字：{str(data)[:500]}")
            return None

    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        print(f"[豆包AI] 调用失败: {e}")
        print(f"[豆包AI] 详细错误:\n{err_detail}")
        raise RuntimeError(f"豆包API调用异常({type(e).__name__}): {e}") from e

    # 过滤无效回答
    if not raw_text or raw_text in ['不确定', '未知', '不清楚', '无法回答', '']:
        return None

    # 提取地址文本（去掉可能的多余文字）
    address = _clean_doubao_address(raw_text)
    if not address:
        return None

    result = {
        'address': address,
        'confidence': 0.92,
        'source': '豆包AI（智能搜索）',
        'raw_response': raw_text,
        'lng': None,
        'lat': None
    }

    # 如果有高德Key，进行地理编码
    if amap_key:
        geo = geocode_text_address(address, amap_key)
        if geo and geo.get('lng'):
            result['lng'] = geo['lng']
            result['lat'] = geo['lat']
            result['geo_source'] = '高德地理编码'

    return result


def _model_supports_search(model_id):
    """判断模型是否支持联网搜索工具（豆包Pro/Plus系列支持）"""
    if not model_id:
        return False
    # 豆包-pro、doubao-pro-32k、doubao-plus 等支持 web_search 工具
    search_supported_keywords = ['pro', 'plus', 'search', 'web']
    model_lower = model_id.lower()
    return any(kw in model_lower for kw in search_supported_keywords)


def _clean_doubao_address(text):
    """清理豆包返回的地址文本"""
    import re
    if not text:
        return None
    # 去掉"申报地区："等前缀
    text = re.sub(r"^(申报地区[：:]\s*|地址[：:]\s*|位于\s*|在\s*)", "", text)
    # 去掉末尾的句号等
    text = re.sub(r"[。，、；;,，\s]+$", "", text)
    # 去掉可能的英文或特殊字符（保留中文地名）
    text = re.sub(r"[a-zA-Z0-9_\-]{3,}", "", text)
    text = text.strip()
    # 长度校验：太少或太多都不是有效地址
    if len(text) < 3 or len(text) > 50:
        return None
    return text



def smart_extract_address(title, snippet='', url=''):
    """
    智能地址提取算法（优化版）
    从标题、摘要、URL等多个来源提取地址信息，并计算置信度
    优先保证地址准确率，宁可不返回地址，也不返回错误地址
    
    返回: {
        'address': 提取到的地址,
        'confidence': 置信度 (0-1),
        'components': 地址组成部分,
        'source': 提取来源 (title/snippet/combined)
    }
    """
    best_address = ''
    best_confidence = 0
    best_components = {}
    best_source = ''
    
    # 构建完整文本（提前构建，供heritage extraction使用）
    full_text = ''
    if title:
        full_text += title
    if snippet:
        if full_text:
            full_text += ' ' + snippet
        else:
            full_text = snippet
    
    # ===== v4新增: 非遗元数据提取优先（模拟AI搜索的结构化分析） =====
    if full_text:
        heritage_result = extract_heritage_info(full_text)
        h_addr = heritage_result.get('address', '')
        h_conf = heritage_result.get('confidence', 0)
        h_src = heritage_result.get('source', '')
        
        if h_addr and h_conf > best_confidence:
            best_address = h_addr
            best_confidence = h_conf
            best_source = f'heritage:{h_src}'
            
            if heritage_result.get('region_detail'):
                best_components['heritage_region'] = heritage_result['region_detail']
            if heritage_result.get('organization'):
                best_components['organization'] = heritage_result['organization']
    
    # 1. 通用地址提取（原有逻辑作为补充/降级方案）
    
    if full_text:
        result = extract_address_with_confidence(full_text)
        if result['confidence'] > best_confidence and result['address']:
            best_address = result['address']
            best_confidence = result['confidence']
            best_components = result['components']
            best_source = 'combined'
    
    # 2. 如果完整文本没提取到，再分别尝试从标题和摘要提取
    if not best_address:
        sources = [
            ('title', title),
            ('snippet', snippet),
        ]
        
        for source_name, text in sources:
            if not text:
                continue
            
            result = extract_address_with_confidence(text)
            # 只有置信度较高的才采用，避免低质量地址
            if result['confidence'] > best_confidence and result['confidence'] >= 0.3 and result['address']:
                best_address = result['address']
                best_confidence = result['confidence']
                best_components = result['components']
                best_source = source_name
    
    # 3. 最终验证：如果地址看起来不合理，直接丢弃
    if best_address:
        # 检查地址是否太短（只有1-2个字）
        if len(best_address) < 2:
            return {'address': '', 'confidence': 0, 'components': {}, 'source': ''}
        
        # 检查地址是否包含明显的非地址字符
        invalid_patterns = ['其中', '其中', '月', '年', '日', '时', '分', '秒', '第', '个']
        for pattern in invalid_patterns:
            if pattern in best_address[:2]:  # 只检查开头
                return {'address': '', 'confidence': 0, 'components': {}, 'source': ''}
    
    return {
        'address': best_address,
        'confidence': best_confidence,
        'components': best_components,
        'source': best_source
    }


def is_valid_place_name(name, min_len=2, max_len=4):
    """
    判断一个字符串是否可能是合理的地名
    过滤掉包含动词、介词等的不合理匹配
    返回: (bool, int) - (是否有效, 质量评分)
    """
    if not name:
        return False, 0
    
    # 长度检查
    if len(name) < min_len or len(name) > max_len:
        return False, 0
    
    score = 50  # 基础分
    
    # 常见的不合理前缀字（动词、介词等）
    invalid_prefixes = ['位', '于', '在', '是', '的', '和', '与', '及', '分', '布', 
                        '发', '源', '起', '坐', '落', '地', '处', '位', '置',
                        '有', '没', '不', '也', '都', '就', '才', '又', '再',
                        '把', '被', '让', '给', '从', '向', '往', '到', '由',
                        '因', '为', '所', '以', '如', '果', '虽', '然', '但',
                        '而', '且', '或', '者', '还', '要', '会', '能', '可',
                        '该', '本', '此', '那', '这', '哪', '每', '各', '某',
                        '东', '西', '南', '北', '中', '上', '下', '左', '右',
                        '前', '后', '里', '外', '内', '旁', '边', '角', '顶',
                        '高', '低', '大', '小', '长', '短', '新', '旧', '老',
                        '第', '初', '末', '始', '终', '首', '尾', '头', '个']
    
    # 检查第一个字是否是不合理的前缀
    if name[0] in invalid_prefixes:
        score -= 30
    
    # 检查是否包含明显的动词
    invalid_chars = ['说', '看', '听', '想', '做', '走', '跑', '跳', '吃', '喝',
                     '写', '读', '学', '习', '工', '作', '生', '活', '打', '玩']
    for char in invalid_chars:
        if char in name:
            score -= 40
    
    # 常见的地名用字（加分）
    common_place_chars = ['州', '市', '县', '区', '镇', '乡', '村', '街', '路',
                          '城', '镇', '堡', '寨', '庄', '屯', '店', '铺', '坊',
                          '山', '水', '河', '湖', '海', '江', '溪', '泉', '潭',
                          '岭', '峰', '岩', '石', '洞', '沟', '谷', '坡', '岗',
                          '门', '口', '湾', '岛', '洲', '岸', '滩', '坝', '塘']
    common_count = 0
    for char in common_place_chars:
        if char in name:
            common_count += 1
    score += min(common_count * 10, 30)
    
    # 长度适中加分（2-3个字最常见）
    if 2 <= len(name) <= 3:
        score += 10
    
    # 判断是否有效
    is_valid = score > 20
    
    return is_valid, score


def extract_clean_place_name(match, suffix):
    """
    从匹配结果中清洗出真正的地名
    例如：从"东省惠州市"中提取出"惠州市"
    
    参数:
        match: 完整的匹配字符串
        suffix: 后缀（如"市"、"县"、"区"等）
    
    返回: 清洗后的地名字符串（包含后缀）
    """
    if not match or not suffix:
        return match
    
    # 去掉后缀
    name_part = match[:-len(suffix)]
    
    if not name_part:
        return match
    
    # 从后往前找，找到最长的合理地名
    # 常见的不合理前缀字
    invalid_prefixes = ['位', '于', '在', '是', '的', '和', '与', '及', '分', '布', 
                        '发', '源', '起', '坐', '落', '地', '处', '位', '置',
                        '有', '没', '不', '也', '都', '就', '才', '又', '再',
                        '把', '被', '让', '给', '从', '向', '往', '到', '由',
                        '因', '为', '所', '以', '如', '果', '虽', '然', '但',
                        '而', '且', '或', '者', '还', '要', '会', '能', '可',
                        '该', '本', '此', '那', '这', '哪', '每', '各', '某',
                        '东', '西', '南', '北', '中', '上', '下', '左', '右',
                        '前', '后', '里', '外', '内', '旁', '边', '角', '顶',
                        '高', '低', '大', '小', '长', '短', '新', '旧', '老',
                        '第', '初', '末', '始', '终', '首', '尾', '头', '个',
                        '省', '市', '县', '区', '镇', '乡', '村', '街', '路']
    
    # 从后往前找，找到第一个不合理的字的位置
    cut_pos = 0
    for i in range(len(name_part) - 1, -1, -1):
        if name_part[i] in invalid_prefixes:
            cut_pos = i + 1
            break
    
    if cut_pos > 0:
        name_part = name_part[cut_pos:]
    
    # 如果清洗后太短，就用原始的
    if len(name_part) < 1:
        return match
    
    return name_part + suffix


def find_best_place_name(matches, min_len=2, max_len=4):
    """
    从多个匹配中选择最佳的地名
    返回最佳匹配的字符串，如果都无效则返回空字符串
    """
    if not matches:
        return ''
    
    best_match = ''
    best_score = 0
    
    for match in matches:
        # 先清洗匹配结果，提取真正的地名
        # 先判断可能的后缀
        suffixes = ['街道', '社区', '省', '市', '区', '县', '镇', '乡', '村', '路', '街', '巷', '弄', '号']
        suffix = ''
        for s in suffixes:
            if match.endswith(s):
                suffix = s
                break
        
        if suffix:
            # 清洗匹配结果
            clean_match = extract_clean_place_name(match, suffix)
        else:
            clean_match = match
        
        # 提取地名部分（只去掉最后一个后缀字）
        name_part = clean_match
        for s in suffixes:
            if name_part.endswith(s):
                name_part = name_part[:-len(s)]
                break
        
        is_valid, score = is_valid_place_name(name_part, min_len, max_len)
        
        if is_valid and score > best_score:
            best_score = score
            best_match = clean_match
    
    return best_match



def extract_heritage_info(text):
    """
    从文本中提取非遗项目的结构化信息（模拟AI搜索的元数据提取）
    优先级最高，因为非遗官网/百科的结构化数据最准确

    返回: dict {
        'address': 提取到的最佳地址,
        'confidence': 置信度(0-1),
        'source': 数据来源字段,
        'region_detail': 详细地区信息,
        'organization': 保护单位/传承基地名称,
        'raw_matches': 原始匹配结果
    }
    """
    if not text:
        return {'address': '', 'confidence': 0, 'source': '', 'region_detail': '', 'organization': '', 'raw_matches': []}

    best_addr = ''
    best_conf = 0
    best_source = ''
    region_detail = ''
    organization = ''
    raw_matches = []

    # ===== 第一优先：申报地区 / 申报地 (最准确) =====
    declaration_patterns = [
        r'申报地区[：:\s]\s*(.{4,40}?)(?:\n|。|；|;|$)',
        r'申报地[点区]?[：:\s]\s*(.{4,40}?)(?:\n|。|；|;|$)',
        r'所属地区[：:\s]\s*(.{4,40}?)(?:\n|。|；|;|$)',
        r'分布地区[：:\s]\s*(.{4,40}?)(?:\n|。|；|;|$)',
        r'所在地[：:\s]\s*(.{4,40}?)(?:\n|。|；|;|$)',
    ]
    for pat in declaration_patterns:
        m = re.search(pat, text)
        if m:
            addr = m.group(1).strip()
            addr = re.sub(r'[<>【】\[\]（）()\s]+$', '', addr)
            addr = re.sub(r'^[<>【】\[\]（）()\s]+', '', addr)
            if len(addr) >= 4 and _looks_like_address(addr):
                best_addr = addr
                best_conf = 0.95
                best_source = '申报地区'
                raw_matches.append(('declaration', addr, 0.95))
                break

    # ===== 第二优先：核心流传区域 / 流传地区 =====
    if not best_addr or best_conf < 0.9:
        region_patterns = [
            r'核心流传区域[：:]?\s*(.{3,50}?)(?:\n|。|；|;|\||代表|$)',
            r'流传地区[：:]?\s*(.{3,50}?)(?:\n|。|；|;|$)',
            r'主要流传[于区]?[：:]?\s*(.{3,50}?)(?:\n|。|；|;|$)',
            r'分布[于]?[：:]?\s*(.{3,50}?)(?:\n|。|；|;|代表|$)',
            r'(?:位于|地处|坐落于)\s*[：:]?\s*(.{3,50}?)(?:\n|。|，|；|;|$)',
        ]
        for pat in region_patterns:
            m = re.search(pat, text)
            if m:
                region = m.group(1).strip()
                region = re.sub(r'[<>【】\[\]（）()]$', '', region.strip())
                if len(region) >= 3 and _looks_like_region(region):
                    region_detail = region
                    # 检查是否包含完整省市区地址
                    if re.search(r'[\u4e00-\u9fa5]{2,3}省.*?(?:市|县|区).*?(?:镇|乡|村|街|路)', region):
                        best_addr = region
                        best_conf = 0.90
                        best_source = '核心流传区域'
                    elif not best_addr and len(region) >= 4:
                        best_addr = region
                        best_conf = 0.80
                        best_source = '流传区域'
                    raw_matches.append(('region', region, 0.85))
                    break

    # ===== 第三优先：保护单位 / 传承基地 =====
    if not best_addr or best_conf < 0.85:
        org_patterns = [
            r'保护单位[：:\s]*(.{2,30}?)(?:地址[：:\s]*)(.{3,40}?)(?:\n|。|；|;|$)',
        ]
        for pat in org_patterns:
            m = re.search(pat, text)
            if m:
                groups = m.groups()
                if len(groups) >= 1 and groups[0]:
                    org_name = groups[0].strip()
                    organization = org_name
                if len(groups) >= 2 and groups[1]:
                    addr = groups[1].strip()
                    addr = re.sub(r'[<>【]\[()\]]+$', '', addr)
                    if len(addr) >= 4 and _looks_like_address(addr):
                        best_addr = addr
                        best_conf = 0.88
                        best_source = f'保护单位({org_name})'
                        raw_matches.append(('org_addr', addr, 0.88))
                        break

    # ===== 第四优先：完整地址正则（省-市-区-镇） =====
    if not best_addr or best_conf < 0.7:
        full_addr_patterns = [
            r'([\u4e00-\u9fa5]{2,3}省[\u4e00-\u9fa5]{2,4}(?:市|自治州)[\u4e00-\u9fa5]{2,4}(?:市|区|县|旗)[\u4e00-\u9fa5]{2,10}(?:镇|乡|村|街道|路))',
            r'([\u4e00-\u9fa5]{2,4}(?:市|自治州)[\u4e00-\u9fa5]{2,4}(?:市|区|县|旗)[\u4e00-\u9fa5]{2,10}(?:镇|乡|村|街道|路))',
        ]
        for pat in full_addr_patterns:
            matches = re.findall(pat, text)
            for addr in matches:
                addr_stripped = addr.strip()
                if _looks_like_address(addr_stripped) and len(addr_stripped) >= 6:
                    best_addr = addr_stripped
                    best_conf = 0.80
                    best_source = '完整地址正则'
                    raw_matches.append(('full_regex', addr_stripped, 0.80))
                    break
            if best_addr and best_conf >= 0.80:
                break

    return {
        'address': best_addr,
        'confidence': best_conf,
        'source': best_source,
        'region_detail': region_detail,
        'organization': organization,
        'raw_matches': raw_matches
    }


def _looks_like_address(text):
    """快速判断文本是否像地址"""
    if not text or len(text) < 3:
        return False
    addr_indicators = ['省', '市', '区', '县', '镇', '乡', '村', '街', '路', '道', '号']
    has_indicator = any(ind in text for ind in addr_indicators)
    if not has_indicator:
        return False
    non_addr_keywords = ['点击', '查看', '更多', '请访问', 'http', 'www.', '登录']
    for kw in non_addr_keywords:
        if kw in text:
            return False
    return True


def _looks_like_region(text):
    """快速判断文本是否像地区描述"""
    if not text or len(text) < 2:
        return False
    indicators = ['省', '市', '区', '县', '镇', '乡', '村', '地区', '流域']
    return any(ind in text for ind in indicators)


def geocode_text_address(address, amap_key):
    """使用高德地理编码API将文本地址转为经纬度坐标"""
    if not address or not amap_key:
        return None
    try:
        url = "https://restapi.amap.com/v3/geocode/geo"
        params = {'key': amap_key, 'address': address, 'output': 'json'}
        resp = requests.get(url, params=params, timeout=10)
        data = resp.json()
        if data.get('status') == '1' and data.get('geocodes'):
            geo = data['geocodes'][0]
            location = geo.get('location', '')
            if location:
                parts = location.split(',')
                lng, lat = parts[0], parts[1] if len(parts) >= 2 else ('', '')
                return {
                    'lng': lng,
                    'lat': lat,
                    'formatted_address': geo.get('formatted_address', address),
                    'province': geo.get('province', ''),
                    'city': geo.get('city', ''),
                    'district': geo.get('district', ''),
                    'level': geo.get('level', ''),
                    'confidence': float(geo.get('confidence', 0)) / 100 if geo.get('confidence') else 0.7
                }
    except Exception as e:
        pass
    return None

def extract_address_with_confidence(text):
    """
    从文本中提取地址并计算置信度（优化版）
    - 修复地址提取不准确的问题（如"月惠州"、"其中惠州客家地区"等）
    - 优化置信度计算（省+市+区=高分，市+区=中分，只有区县=低分）
    - 置信度≥0.7自动通过
    """
    if not text:
        return {'address': '', 'confidence': 0, 'components': {}}
    
    components = {}
    address_parts = []
    
    # 1. 先尝试匹配显式格式（地址：xxx、位于：xxx等）
    explicit_patterns = [
        r'地址[：:]\s*([^，。；;\n]+)',
        r'位于[：:]\s*([^，。；;\n]+)',
        r'坐落于\s*([^，。；;\n]+)',
        r'地处\s*([^，。；;\n]+)',
        r'地址为\s*([^，。；;\n]+)',
        r'分布于\s*([^，。；;\n]+)',
        r'发源地\s*([^，。；;\n]+)',
        r'起源于\s*([^，。；;\n]+)',
    ]
    
    explicit_addr = ''
    for pattern in explicit_patterns:
        match = re.search(pattern, text)
        if match:
            addr = match.group(1).strip()
            # 清理多余字符
            addr = re.sub(r'[<>【】\[\]（）()]', '', addr)
            if len(addr) > 3 and len(addr) < 100:
                explicit_addr = addr
                break
    
    # 如果有显式地址，优先从显式地址中提取组件
    target_text = explicit_addr if explicit_addr else text
    
    # 2. 提取所有可能的地址元素
    # 使用 findall 找到所有匹配，然后选择最合理的
    
    # 省级行政区（精确匹配优先）
    provinces = [
        '北京市', '天津市', '上海市', '重庆市',
        '河北省', '山西省', '辽宁省', '吉林省', '黑龙江省',
        '江苏省', '浙江省', '安徽省', '福建省', '江西省', '山东省',
        '河南省', '湖北省', '湖南省', '广东省', '海南省',
        '四川省', '贵州省', '云南省', '陕西省', '甘肃省', '青海省',
        '台湾省',
        '内蒙古自治区', '广西壮族自治区', '西藏自治区', '宁夏回族自治区', '新疆维吾尔自治区',
        '香港特别行政区', '澳门特别行政区'
    ]
    
    # 先尝试精确匹配省份
    for prov in provinces:
        if prov in target_text:
            components['province'] = prov
            address_parts.append(prov)
            break
    
    # 如果没精确匹配到，尝试用正则匹配
    if 'province' not in components:
        # 匹配"XX省"（2-3个字+省）
        prov_matches = re.findall(r'[\u4e00-\u9fa5]{2,3}省', target_text)
        if prov_matches:
            # 从多个匹配中选择最佳的
            best_prov = find_best_place_name(prov_matches, min_len=2, max_len=3)
            if best_prov:
                # 额外验证：省名前面不能是动词或介词
                valid_prov = True
                prov_pos = target_text.find(best_prov)
                if prov_pos > 0:
                    prev_char = target_text[prov_pos - 1]
                    invalid_prev = ['在', '于', '是', '的', '和', '与', '及', '分', '布',
                                    '发', '源', '起', '坐', '落', '地', '处', '位', '置',
                                    '有', '没', '不', '也', '都', '就', '才', '又', '再',
                                    '把', '被', '让', '给', '从', '向', '往', '到', '由',
                                    '因', '为', '所', '以', '如', '果', '虽', '然', '但',
                                    '而', '且', '或', '者', '还', '要', '会', '能', '可',
                                    '该', '本', '此', '那', '这', '哪', '每', '各', '某',
                                    '东', '西', '南', '北', '中', '上', '下', '左', '右',
                                    '前', '后', '里', '外', '内', '旁', '边', '角', '顶',
                                    '高', '低', '大', '小', '长', '短', '新', '旧', '老',
                                    '第', '初', '末', '始', '终', '首', '尾', '头', '个',
                                    '月', '年', '日', '时', '分', '秒']
                    if prev_char in invalid_prev:
                        valid_prov = False  # 这个匹配可能是错误的，跳过
                
                if valid_prov:
                    components['province'] = best_prov
                    address_parts.append(best_prov)
    
    # 3. 市级行政区
    # 匹配"XX市"、"XX州"、"XX盟"
    city_matches = re.findall(r'[\u4e00-\u9fa5]{2,4}市', target_text)
    state_matches = re.findall(r'[\u4e00-\u9fa5]{2,3}州', target_text)
    league_matches = re.findall(r'[\u4e00-\u9fa5]{2,3}盟', target_text)
    
    all_city_matches = city_matches + state_matches + league_matches
    
    # 从多个匹配中选择最佳的
    best_city = find_best_place_name(all_city_matches, min_len=2, max_len=4)
    if best_city:
        # 避免和省份重复
        if not ('province' in components and best_city == components['province']):
            # 避免城市名包含"省"字
            if '省' not in best_city:
                # 额外验证：市名前面不能是明显的动词或介词
                city_pos = target_text.find(best_city)
                if city_pos > 0:
                    prev_char = target_text[city_pos - 1]
                    invalid_prev = ['在', '于', '是', '的', '和', '与', '及', '分', '布',
                                    '发', '源', '起', '坐', '落', '地', '处', '位', '置',
                                    '有', '没', '不', '也', '都', '就', '才', '又', '再',
                                    '把', '被', '让', '给', '从', '向', '往', '到', '由',
                                    '因', '为', '所', '以', '如', '果', '虽', '然', '但',
                                    '而', '且', '或', '者', '还', '要', '会', '能', '可',
                                    '该', '本', '此', '那', '这', '哪', '每', '各', '某',
                                    '东', '西', '南', '北', '中', '上', '下', '左', '右',
                                    '前', '后', '里', '外', '内', '旁', '边', '角', '顶',
                                    '高', '低', '大', '小', '长', '短', '新', '旧', '老',
                                    '第', '初', '末', '始', '终', '首', '尾', '头', '个',
                                    '月', '年', '日', '时', '分', '秒', '其', '中', '其']
                    if prev_char in invalid_prev:
                        # 检查前面是否有"省"字，如果有则是合理的（如"广东省惠州市"）
                        has_province_before = False
                        if 'province' in components:
                            prov_pos = target_text.find(components['province'])
                            if prov_pos >= 0 and prov_pos < city_pos:
                                has_province_before = True
                        
                        if not has_province_before:
                            # 可能是错误的匹配，比如"其中惠州..."，跳过
                            pass
                        else:
                            components['city'] = best_city
                            address_parts.append(best_city)
                    else:
                        components['city'] = best_city
                        address_parts.append(best_city)
                else:
                    components['city'] = best_city
                    address_parts.append(best_city)
    
    # 4. 区/县级
    district_matches = re.findall(r'[\u4e00-\u9fa5]{2,4}区', target_text)
    county_matches = re.findall(r'[\u4e00-\u9fa5]{2,4}县', target_text)
    
    all_district_matches = district_matches + county_matches
    
    # 过滤掉明显不合理的区名（如景区、度假区、开发区等）
    invalid_district_keywords = ["景区", "度假", "风景", "开发", "工业", "农业", "商业", "科技", "经济", "贸易", "物流", "产业", "示范", "实验", "生态", "文化", "旅游"]
    filtered_district_matches = []
    for match in all_district_matches:
        is_valid = True
        for kw in invalid_district_keywords:
            if kw in match:
                is_valid = False
                break
        if is_valid:
            filtered_district_matches.append(match)
    
    # 从多个匹配中选择最佳的
    best_district = find_best_place_name(filtered_district_matches, min_len=2, max_len=4)
    if best_district:
        # 避免和市级重复
        if not ('city' in components and best_district == components['city']):
            # 避免和省级重复
            if not ('province' in components and best_district == components['province']):
                # 避免包含"市"或"省"字
                if '市' not in best_district and '省' not in best_district:
                    components['district'] = best_district
                    address_parts.append(best_district)
    
    # 5. 乡镇/街道
    town_matches = re.findall(r'[\u4e00-\u9fa5]{2,5}镇', target_text)
    township_matches = re.findall(r'[\u4e00-\u9fa5]{2,5}乡', target_text)
    street_matches = re.findall(r'[\u4e00-\u9fa5]{2,5}街道', target_text)
    
    all_town_matches = town_matches + township_matches + street_matches
    
    # 从多个匹配中选择最佳的
    best_town = find_best_place_name(all_town_matches, min_len=2, max_len=5)
    if best_town:
        # 避免包含"区"、"县"、"市"、"省"字
        if not any(x in best_town for x in ['区', '县', '市', '省']):
            components['town'] = best_town
            address_parts.append(best_town)
    
    # 6. 详细地址（路/街/巷/号/弄/村）
    detail_patterns = [
        r'[\u4e00-\u9fa5\d]{2,20}路\d+号',
        r'[\u4e00-\u9fa5\d]{2,20}街\d+号',
        r'[\u4e00-\u9fa5\d]{2,20}巷\d+号',
        r'[\u4e00-\u9fa5\d]{2,20}弄\d+号',
    ]
    
    for pattern in detail_patterns:
        matches = re.findall(pattern, target_text)
        if matches:
            # 选择第一个
            detail = matches[0]
            # 避免包含"镇"、"乡"、"街道"等
            if not any(x in detail for x in ['镇', '乡', '街道', '区', '县']):
                components['detail'] = detail
                address_parts.append(detail)
                break
    
    # 村/社区
    if 'detail' not in components:
        village_matches = re.findall(r'[\u4e00-\u9fa5]{2,10}村', target_text)
        community_matches = re.findall(r'[\u4e00-\u9fa5]{2,10}社区', target_text)
        all_village_matches = village_matches + community_matches
        
        for village in all_village_matches:
            if not any(x in village for x in ['镇', '乡', '街道', '区', '县', '市', '省']):
                name_part = re.sub(r'[村社区]', '', village)
                if is_valid_place_name(name_part, min_len=2, max_len=10):
                    components['detail'] = village
                    address_parts.append(village)
                    break
    
    # 7. 如果有显式地址但没提取到组件，直接使用显式地址
    if explicit_addr and not address_parts:
        components['explicit'] = explicit_addr
        address_parts.append(explicit_addr)
    
    # 计算置信度（优化版）
    confidence = 0
    
    # 有明确的"地址："格式，基础分高
    if 'explicit' in components or explicit_addr:
        confidence += 0.25
    
    # 根据地址组成部分计算置信度（优化权重）
    if 'province' in components and 'city' in components and 'district' in components:
        confidence += 0.55  # 省+市+区 = 高分（0.55）
    elif 'province' in components and 'city' in components:
        confidence += 0.45  # 省+市 = 中高分
    elif 'city' in components and 'district' in components:
        confidence += 0.45  # 市+区 = 中高分
    elif 'province' in components and 'district' in components:
        confidence += 0.35  # 省+区 = 中分
    elif 'city' in components and 'town' in components:
        confidence += 0.35  # 市+镇 = 中分
    elif 'district' in components and 'town' in components:
        confidence += 0.3   # 区+镇 = 中低分
    elif 'city' in components:
        confidence += 0.25  # 只有市 = 低分
    elif 'district' in components:
        confidence += 0.2   # 只有区 = 更低分
    elif 'town' in components:
        confidence += 0.15  # 只有镇 = 最低分
    
    # 有详细地址加分
    if 'detail' in components:
        confidence += 0.15
    
    # 地址长度合理性
    if address_parts:
        full_address = ''.join(address_parts)
        if 5 <= len(full_address) <= 50:
            confidence += 0.05
        elif len(full_address) > 50:
            confidence -= 0.1
    
    # 确保置信度在0-1之间
    confidence = max(0, min(1, confidence))
    
    # 组装完整地址
    full_address = ''
    if address_parts:
        # 按照从大到小的顺序组装
        order = ['province', 'city', 'district', 'town', 'detail', 'explicit']
        ordered_parts = []
        for key in order:
            if key in components:
                ordered_parts.append(components[key])
        full_address = ''.join(ordered_parts)
    
    # 额外验证：如果地址看起来不合理，降低置信度或直接丢弃
    if full_address:
        # 检查是否有重复的字（比如"广东省省"）
        if re.search(r'(.)\1{2,}', full_address):
            confidence -= 0.2
        
        # 检查是否以奇怪的字开头
        if full_address[0] in '是于在的和与及其中':
            confidence -= 0.2
            # 如果开头是"其中"等，直接丢弃这个地址
            if full_address.startswith('其中') or full_address.startswith('其'):
                return {'address': '', 'confidence': 0, 'components': {}}
        
        # 检查是否包含"省"字在错误位置（比如"东省惠州市"，省字前面只有一个字）
        if '省' in full_address:
            sheng_pos = full_address.index('省')
            if sheng_pos == 1:  # 省字前面只有一个字（如"X省..."），可能是错误的
                confidence -= 0.2
                # 如果省前面只有一个字，且不是直辖市的特殊情况，直接丢弃
                if len(full_address) > 3:
                    return {'address': '', 'confidence': 0, 'components': {}}
        
        # 检查地址是否太短（只有2-3个字且没有后缀）
        if len(full_address) <= 3:
            confidence -= 0.1
    
    confidence = max(0, min(1, confidence))
    
    return {
        'address': full_address,
        'confidence': confidence,
        'components': components
    }


def is_valid_address(address, min_confidence=0.3):
    """
    判断地址是否有效
    """
    if not address:
        return False
    
    result = extract_address_with_confidence(address)
    return result['confidence'] >= min_confidence


def extract_place_name(keyword):
    """
    从项目名称中提取地名（优化版）
    优先提取明确的地名（带市/县/区等后缀），其次尝试提取常见地名
    确保提取的是真实的地名，不是随便的文字
    """
    if not keyword:
        return ''
    
    # 1. 优先提取包含明确后缀的地名（市/县/区/镇/乡）
    suffixes = ['市', '县', '区', '镇', '乡']
    
    for suffix in suffixes:
        pattern = r'([\u4e00-\u9fa5]{2,4}' + suffix + ')'
        match = re.search(pattern, keyword)
        if match:
            place = match.group(1)
            # 验证：地名前面不能是明显的非地名字
            if len(place) > 2:
                return place
    
    # 2. 尝试匹配常见的县级及以上地名（简化版，只匹配常见的）
    # 常见的单字县名（如"龙门"、"博罗"等）
    # 这里简单处理：如果项目名前两个字是常见地名格式，就提取
    if len(keyword) >= 2:
        # 检查前两个字是否可能是地名
        first_two = keyword[:2]
        # 简单验证：不是常见的非地名词
        invalid_prefixes = ['制作', '传统', '技艺', '工艺', '文化', '民俗', '非遗', '传承', '遗产', '项目']
        if first_two not in invalid_prefixes:
            # 再检查第三个字，如果是"米"、"饼"、"茶"等，说明前两个字可能是地名
            if len(keyword) >= 3:
                third_char = keyword[2]
                # 如果第三个字是常见的非地名字，说明前两个字可能是地名
                if third_char in ['米', '饼', '茶', '酒', '醋', '糖', '面', '粉', '菜', '药', '绣', '雕', '画', '纸', '布', '丝', '竹', '木', '石', '陶', '瓷']:
                    return first_two
    
    # 3. 如果还是没找到，返回空字符串（不瞎猜）
    return ''

def extract_address_from_text(text):
    """从文本中提取地址信息"""
    # 常见地址模式
    patterns = [
        r'地址[：:]\s*([^，。；;\n]+)',
        r'位于[：:]\s*([^，。；;\n]+)',
        r'坐落于\s*([^，。；;\n]+)',
        r'地处\s*([^，。；;\n]+)',
        r'地址为\s*([^，。；;\n]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            address = match.group(1).strip()
            # 清理多余字符
            address = re.sub(r'[<>【】\[\]]', '', address)
            if len(address) > 5 and len(address) < 100:
                return address
    
    return ''

def get_source_name(url):
    """根据URL推断来源网站名称"""
    if not url:
        return '未知来源'
    
    try:
        from urllib.parse import urlparse
        parsed = urlparse(url)
        domain = parsed.netloc.lower()
        
        # 常见网站映射
        source_map = {
            'baike.baidu.com': '百度百科',
            'baike.sogou.com': '搜狗百科',
            'baike.so.com': '360百科',
            'www.gov.cn': '中国政府网',
            'www.mct.gov.cn': '文化和旅游部',
            'www.ihchina.cn': '中国非物质文化遗产网',
            'www.gd.gov.cn': '广东省人民政府网',
            'www.huizhou.gov.cn': '惠州市人民政府网',
            'www.longmen.gov.cn': '龙门县人民政府网',
            'www.people.com.cn': '人民网',
            'www.xinhuanet.com': '新华网',
            'www.chinadaily.com.cn': '中国日报网',
            'www.baidu.com': '百度',
            'www.sohu.com': '搜狐网',
            'www.sina.com.cn': '新浪网',
            'www.163.com': '网易',
            'www.qq.com': '腾讯网',
            'baike.zhshch.com': '中华非遗网',
        }
        
        # 精确匹配
        if domain in source_map:
            return source_map[domain]
        
        # 模糊匹配
        if 'gov.cn' in domain:
            return '政府网站'
        elif 'baike' in domain:
            return '百科网站'
        elif 'people' in domain or 'xinhua' in domain:
            return '官方媒体'
        elif '非遗' in domain or 'heritage' in domain or 'feiyi' in domain:
            return '非遗相关网站'
        elif 'edu.cn' in domain:
            return '教育机构'
        else:
            # 提取主域名
            parts = domain.split('.')
            if len(parts) >= 2:
                main_domain = parts[-2] + '.' + parts[-1]
                return main_domain
            return domain
            
    except Exception:
        return '未知来源'

def get_source_type(source_name, url=''):
    """根据来源名称或URL判断来源类型"""
    source_lower = source_name.lower()
    url_lower = url.lower()
    
    # 政府官网
    if '政府' in source_name or 'gov.cn' in url_lower or '人民政府' in source_name:
        return 'government'
    # 百科类
    elif '百科' in source_name or 'baike' in url_lower:
        return 'encyclopedia'
    # 官方媒体
    elif '人民' in source_name or '新华' in source_name or '日报' in source_name or '官方媒体' in source_name:
        return 'media'
    # 非遗相关
    elif '非遗' in source_name or '非物质文化遗产' in source_name or '文化和旅游' in source_name or 'heritage' in source_lower or 'ihchina' in url_lower or 'feiyi' in url_lower:
        return 'heritage'
    # 其他
    else:
        return 'other'

def sort_search_results(results):
    """
    对搜索结果进行排序，优先展示权威来源
    优先级：政府官网 > 非遗相关 > 百科类 > 官方媒体 > 其他
    """
    # 定义优先级
    priority = {
        'government': 1,    # 政府官网
        'heritage': 2,      # 非遗相关
        'encyclopedia': 3,  # 百科类
        'media': 4,         # 官方媒体
        'other': 5          # 其他
    }
    
    # 按优先级排序
    sorted_results = sorted(results, key=lambda x: priority.get(x.get('source_type', 'other'), 5))
    
    # 重新编号
    for idx, result in enumerate(sorted_results):
        result['index'] = idx + 1
    
    return sorted_results


def multi_level_search(keyword, min_confidence=0.6, provinces=None, amap_key=None, doubao_api_key=None, doubao_model_id=None):
    """
    多级关键词搜索策略（简化优化版）
    参数:
        keyword: 项目名称
        min_confidence: 最小置信度阈值
        provinces: 省份列表（可选）
        amap_key: 高德API Key（可选），用于高德地点搜索
    # 按3级关键词顺序尝试搜索，每级同时尝试百度和必应，找到有效地址就停止    # 优先保证地址准确率，宁可不返回地址，也不返回错误地址    
    # 参数:        keyword: 项目名称
        min_confidence: 最低置信度阈值（默认0.6，中等置信度以上才采用）
        provinces: 省份列表（可选），限定搜索省份以提高准确度
    
    # 返回:        {
            'success': bool,
            'results': 搜索结果列表,
            'best_address': 最佳地址,
            'best_confidence': 最佳置信度,
            'level': 使用的关键词级别 (1-3),
            'level_keyword': 使用的关键词,
            'search_log': 搜索日志列表,
            'place_name': 提取的地名,
            'result_status': 结果状态 (auto_approved/pending_review/manual_needed)
        }
    """
    # 如果有省份限定，生成省份简称字符串
    province_str = ''
    if provinces:
        province_str = ' '.join([p.replace('省','').replace('市','').replace('自治区','')
                                  .replace('壮族','').replace('回族','').replace('维吾尔','')
                                  .replace('特别行政区','') for p in provinces[:3]])
    
    # 3级关键词策略 - 从精确到宽泛逐步降级（简化版，避免过度复杂）
    if province_str:
        keyword_levels = [
            (1, f"{keyword} {province_str} 地址", f"项目名+{province_str}+地址（最直接，优先保证准确率）"),
            (2, f"{keyword} {province_str} 非遗 在哪里", f"项目名+{province_str}+非遗+在哪里（稍宽泛）"),
            (3, f"{keyword} 分布地区", "项目名+分布地区（最宽泛，兜底）"),
        ]
    else:
        keyword_levels = [
            (1, f"{keyword} 地址", "项目名+地址（最直接，优先保证准确率）"),
            (2, f"{keyword} 非遗 在哪里", "项目名+非遗+在哪里（稍宽泛）"),
            (3, f"{keyword} 分布地区", "项目名+分布地区（最宽泛，兜底）"),
        ]
    
    search_log = []
    best_results = []
    best_address = ''
    best_confidence = 0
    best_level = 0
    best_level_keyword = ''
    best_result_item = None
    
    # 提取项目中的地名，用于后续匹配和过滤
    place_name = extract_place_name(keyword)
    
    for level, level_keyword, level_desc in keyword_levels:
        log_entry = {
            'level': level,
            'keyword': level_keyword,
            'description': level_desc,
            'baidu_result_count': 0,
            'bing_result_count': 0,
            'found_valid': False,
            'best_address': '',
            'best_confidence': 0,
            'fail_reason': ''
        }
        
        level_all_results = []
        
        # 1. 先尝试百度搜索
        try:
            baidu_results = search_heritage_address(keyword, custom_keyword=level_keyword, provinces=provinces, amap_key=amap_key, doubao_api_key=doubao_api_key, doubao_model_id=doubao_model_id)
            log_entry['baidu_result_count'] = len(baidu_results) if baidu_results else 0
            if baidu_results:
                level_all_results.extend(baidu_results)
        except Exception as e:
            log_entry['baidu_error'] = str(e)
        
        # 2. 再尝试必应搜索（补充结果）
        try:
            bing_results = fallback_search(keyword, custom_keyword=level_keyword, provinces=provinces, amap_key=amap_key)
            log_entry['bing_result_count'] = len(bing_results) if bing_results else 0
            if bing_results:
                level_all_results.extend(bing_results)
        except Exception as e:
            log_entry['bing_error'] = str(e)
        
        if not level_all_results:
            log_entry['fail_reason'] = '无搜索结果'
            search_log.append(log_entry)
            continue
        
        # 3. 对结果进行过滤和排序
        filtered_results = filter_and_sort_results(level_all_results, keyword, place_name)
        
        if not filtered_results:
            log_entry['fail_reason'] = '过滤后无有效结果'
            search_log.append(log_entry)
            continue
        
        # 4. 对每个结果进行智能地址提取和评分
        level_best_address = ''
        level_best_confidence = 0
        level_best_result = None
        
        for result in filtered_results:
            # 智能提取地址
            extract_result = smart_extract_address(
                result.get('title', ''),
                result.get('snippet', ''),
                result.get('source_url', '')
            )
            
            # 保存提取结果
            result['smart_address'] = extract_result['address']
            result['address_confidence'] = extract_result['confidence']
            result['address_components'] = extract_result['components']
            result['address_source'] = extract_result['source']
            
            # 地名匹配强加分（如果项目名包含地名，结果地址也包含该地名，大幅加分）
            if place_name and extract_result['address']:
                if place_name in extract_result['address']:
                    result['address_confidence'] = min(1.0, result['address_confidence'] + 0.15)
                # 标题或摘要包含地名也加分
                elif place_name in (result.get('title', '') + result.get('snippet', '')):
                    result['address_confidence'] = min(1.0, result['address_confidence'] + 0.08)
            
            # 权威来源地址置信度加分
            source_type = result.get('source_type', 'other')
            if source_type == 'government':
                result['address_confidence'] = min(1.0, result['address_confidence'] + 0.1)
            elif source_type == 'encyclopedia':
                result['address_confidence'] = min(1.0, result['address_confidence'] + 0.05)
            
            # 更新最佳结果
            if result['address_confidence'] > level_best_confidence and extract_result['address']:
                level_best_confidence = result['address_confidence']
                level_best_address = extract_result['address']
                level_best_result = result
        
        log_entry['best_address'] = level_best_address
        log_entry['best_confidence'] = level_best_confidence
        
        # 检查是否找到高置信度有效地址
        if level_best_confidence >= min_confidence and level_best_address:
            log_entry['found_valid'] = True
            search_log.append(log_entry)
            
            # 找到有效地址，返回结果
            best_results = filtered_results
            best_address = level_best_address
            best_confidence = level_best_confidence
            best_level = level
            best_level_keyword = level_keyword
            best_result_item = level_best_result
            break
        
        # 记录失败原因
        if not level_best_address:
            log_entry['fail_reason'] = '未能提取到有效地址'
        elif level_best_confidence < min_confidence:
            log_entry['fail_reason'] = f'置信度不足({level_best_confidence:.2f} < {min_confidence})'
        
        search_log.append(log_entry)
    
    # 确定结果状态
    if best_confidence >= 0.8 and best_address:
        result_status = 'auto_approved'  # 高置信度自动确认
    elif best_confidence >= 0.5 and best_address:
        result_status = 'pending_review'  # 中等置信度待审核
    else:
        result_status = 'manual_needed'  # 低置信度需人工处理
    
    success = best_confidence >= 0.5 and best_address  # 只要有中等以上置信度就算成功找到
    
    return {
        'success': success,
        'results': best_results,
        'best_address': best_address,
        'best_confidence': best_confidence,
        'best_result': best_result_item,
        'level': best_level,
        'level_keyword': best_level_keyword,
        'search_log': search_log,
        'place_name': place_name,
        'result_status': result_status,
        'total_levels_tried': len(search_log)
    }


def calculate_relevance(result, keyword, place_name=''):
    """
    # 计算搜索结果的相关性评分    # 分数越高越相关    """
    score = 0
    title = (result.get('title', '') or '')
    snippet = (result.get('snippet', '') or '')
    full_text = title + ' ' + snippet
    full_text_lower = full_text.lower()
    keyword_lower = keyword.lower()
    
    # 1. 关键词匹配（核心）
    # 标题完全包含关键词（加分最多）
    if keyword in title:
        score += 60
    # 标题包含部分关键词
    elif keyword_lower in title.lower():
        score += 40
    # 摘要包含关键词
    if keyword in snippet:
        score += 30
    elif keyword_lower in snippet.lower():
        score += 20
    
    # 2. 地名匹配（如果项目名包含地名）- 大幅增加权重
    if place_name and len(place_name) >= 2:
        place_in_title = place_name in title
        place_in_snippet = place_name in snippet
        place_in_address = False
        
        # 检查提取的地址中是否包含地名
        smart_addr = result.get('smart_address', '') or result.get('extracted_address', '')
        if smart_addr and place_name in smart_addr:
            place_in_address = True
        
        if place_in_title:
            score += 50  # 标题包含地名，大幅加分
        elif place_in_address:
            score += 40  # 地址包含地名，大幅加分
        elif place_in_snippet:
            score += 30  # 摘要包含地名，加分
    
    # 3. 非遗相关词汇
    heritage_keywords = ['非遗', '非物质文化遗产', '传承', '传统', '文化遗产', '传承基地', '保护中心']
    heritage_count = 0
    for kw in heritage_keywords:
        if kw in full_text:
            heritage_count += 1
    score += min(heritage_count * 10, 30)  # 最多加30分
    
    # 4. 地址相关词汇
    address_keywords = ['地址', '在哪里', '位置', '位于', '坐落', '分布', '发源地', '起源']
    address_count = 0
    for kw in address_keywords:
        if kw in full_text:
            address_count += 1
    score += min(address_count * 8, 25)  # 最多加25分
    
    # 5. 智能地址提取加分
    if result.get('smart_address') and result.get('address_confidence', 0) > 0:
        score += result['address_confidence'] * 40  # 最多加40分
    elif result.get('extracted_address'):
        score += 20
    
    # 6. 权威来源加分
    source_type = result.get('source_type', 'other')
    source_scores = {
        'government': 35,    # 政府官网
        'heritage': 30,      # 非遗相关
        'encyclopedia': 25,  # 百科类
        'media': 15,         # 官方媒体
        'other': 0           # 其他
    }
    score += source_scores.get(source_type, 0)
    
    # 7. 广告过滤（严重减分）
    ad_keywords = ['广告', '推广', '赞助商', 'sponsored', '广告位', '招商', '加盟']
    for kw in ad_keywords:
        if kw in title:
            score -= 150
            break
    
    # 8. 无关内容过滤（减分）
    unrelated_keywords = ['招聘', '求职', '二手', '租房', '买房', '外卖', '快递']
    for kw in unrelated_keywords:
        if kw in title:
            score -= 50
            break
    
    return score

def filter_and_sort_results(results, keyword, place_name=''):
    """
    # 过滤和排序搜索结果（优化版）    - 严格的关键词过滤（必须包含项目名核心词）
    - 广告过滤
    - 权威来源优先排序
    - 地名匹配优先
    """
    if not results:
        return []
    
    # 1. 计算相关性评分
    for result in results:
        result['relevance_score'] = calculate_relevance(result, keyword, place_name)
    
    # 2. 严格的关键词过滤：标题或摘要必须包含项目名称的核心部分
    # 提取项目名的核心词（去掉"制作技艺"、"传统"等后缀）
    core_keywords = extract_core_keywords(keyword)
    
    keyword_filtered = []
    for result in results:
        title = (result.get('title', '') or '')
        snippet = (result.get('snippet', '') or '')
        full_text = title + snippet
        
        # 必须至少包含一个核心关键词
        has_core_keyword = False
        for kw in core_keywords:
            if len(kw) >= 2 and kw in full_text:
                has_core_keyword = True
                break
        
        # 如果有地名，地名也必须出现在结果中（强约束）
        if place_name and len(place_name) >= 2:
            if place_name not in title and place_name not in snippet:
                # 地名不在标题或摘要中，降低优先级但不直接过滤
                result['relevance_score'] = result.get('relevance_score', 0) - 30
        
        if has_core_keyword:
            keyword_filtered.append(result)
    
    # 如果过滤后结果太少，适当放宽（但至少要包含前2个字）
    if len(keyword_filtered) < 2:
        min_keyword = keyword[:2] if len(keyword) >= 2 else keyword
        for result in results:
            if result not in keyword_filtered:
                title = (result.get('title', '') or '')
                snippet = (result.get('snippet', '') or '')
                if min_keyword in title or min_keyword in snippet:
                    keyword_filtered.append(result)
    
    # 如果还是太少，保留前5条
    if len(keyword_filtered) < 2:
        keyword_filtered = results[:5]
    
    # 3. 广告和垃圾内容过滤（严重减分或直接过滤）
    filtered = []
    for result in keyword_filtered:
        title = (result.get('title', '') or '')
        snippet = (result.get('snippet', '') or '')
        full_text = title + snippet
        
        # 广告关键词检测
        ad_keywords = ['广告', '推广', '赞助商', 'sponsored', '广告位', '招商', '加盟', 
                       '代理', '培训', '教程', '下载', '游戏', '娱乐']
        ad_count = sum(1 for kw in ad_keywords if kw in title)
        
        if ad_count >= 2:
            # 明显的广告，直接过滤
            continue
        
        # 无关内容检测
        unrelated_keywords = ['招聘', '求职', '二手', '租房', '买房', '外卖', '快递', 
                              '小说', '电影', '音乐', '视频']
        unrelated_count = sum(1 for kw in unrelated_keywords if kw in title)
        
        if unrelated_count >= 2:
            # 明显无关，直接过滤
            continue
        
        # 相关性分数太低的过滤
        if result.get('relevance_score', 0) < 0:
            continue
            
        filtered.append(result)
    
    # 如果过滤后太少，就保留前5条
    if len(filtered) < 2:
        filtered = keyword_filtered[:5]
    
    # 4. 去重处理（相似地址合并 + 相似标题合并）
    deduped = []
    seen_addresses = set()
    seen_titles = set()
    
    for result in filtered:
        # 标题去重
        title = (result.get('title', '') or '')
        title_key = title[:10] if len(title) >= 10 else title
        if title_key in seen_titles:
            continue
        seen_titles.add(title_key)
        
        # 地址去重
        addr = result.get('smart_address', '') or result.get('extracted_address', '')
        if addr:
            addr_key = addr[:8] if len(addr) >= 8 else addr
            if addr_key in seen_addresses:
                continue
            seen_addresses.add(addr_key)
        
        deduped.append(result)
    
    # 5. 按相关性评分排序（高到低）
    sorted_results = sorted(deduped, key=lambda x: x.get('relevance_score', 0), reverse=True)
    
    # 只保留前8条结果
    sorted_results = sorted_results[:8]
    
    # 重新编号
    for idx, result in enumerate(sorted_results):
        result['index'] = idx + 1
    
    return sorted_results


def extract_core_keywords(keyword):
    """
    # 从项目名称中提取核心关键词    # 用于搜索结果的关键词过滤    """
    if not keyword:
        return []
    
    keywords = [keyword]  # 完整名称
    
    # 去掉常见后缀
    suffixes = ['制作技艺', '传统技艺', '技艺', '制作工艺', '工艺', '艺术', '文化',
                '民俗', '传统', '非遗', '项目', '传承', '遗产']
    
    cleaned = keyword
    for suffix in suffixes:
        if cleaned.endswith(suffix):
            cleaned = cleaned[:-len(suffix)]
            if cleaned:
                keywords.append(cleaned)
    
    # 如果有地名，提取地名+核心词
    # 简单处理：前2-4个字可能是地名
    if len(keyword) >= 4:
        keywords.append(keyword[:2])
        keywords.append(keyword[:3])
        keywords.append(keyword[:4])
    
    # 去重并按长度排序（长的优先）
    unique_keywords = list(set(keywords))
    unique_keywords.sort(key=lambda x: len(x), reverse=True)
    
    return unique_keywords


# ==================== 高德地图API功能 ====================
def geocode_address(address, amap_key, provinces=None):
    """
    # 使用高德地图API将地址转换为坐标    # 返回坐标结果列表    # 参数:        address: 地址字符串
        amap_key: 高德Web服务API Key
        provinces: 省份列表（可选），当只有1个省份时传给高德city参数限定搜索范围
    """
    results = []
    try:
        url = 'https://restapi.amap.com/v3/geocode/geo'
        params = {
            'key': amap_key,
            'address': address,
            'output': 'json'
        }
        # 如果只限定了1个省份，传给高德city参数（多省份时不限定，避免遗漏）
        if provinces and len(provinces) == 1:
            params['city'] = provinces[0]
        
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        
        if data.get('status') == '1' and data.get('geocodes'):
            for idx, geo in enumerate(data['geocodes'][:5]):
                location = geo.get('location', '')
                if location:
                    lng, lat = location.split(',')
                    results.append({
                        'index': idx + 1,
                        'formatted_address': geo.get('formatted_address', ''),
                        'province': geo.get('province', ''),
                        'city': geo.get('city', ''),
                        'district': geo.get('district', ''),
                        'longitude': float(lng),
                        'latitude': float(lat),
                        'level': geo.get('level', ''),
                        'source': '高德地图'
                    })
        else:
            results.append({
                'index': 1,
                'formatted_address': f"地理编码失败: {data.get('info', '未知错误')}",
                'province': '',
                'city': '',
                'district': '',
                'longitude': 0,
                'latitude': 0,
                'level': '',
                'source': '错误'
            })
            
    except Exception as e:
        results.append({
            'index': 1,
            'formatted_address': f'请求出错: {str(e)}',
            'province': '',
            'city': '',
            'district': '',
            'longitude': 0,
            'latitude': 0,
            'level': '',
            'source': '错误'
        })
    
    return results


# ==================== 工作流引擎 ====================
def process_task(task_id):
    """
    # 后台任务处理线程    # 按照工作流逐步处理数据    """
    try:
        task = get_task(task_id)
        if not task:
            return
        
        # 检查是否需要暂停或中断
        def check_pause_interrupt():
            try:
                task = get_task(task_id)
                if not task:
                    return True
                
                if task.get('interrupted'):
                    try:
                        add_log(task_id, '任务已被中断', 'warning')
                        update_task(task_id, status=TASK_STATUS['INTERRUPTED'])
                    except:
                        pass  # 确保中断时日志记录失败不会影响中断流程
                    return True
                
                if task.get('paused'):
                    try:
                        add_log(task_id, '任务已暂停', 'info')
                        update_task(task_id, status=TASK_STATUS['PAUSED'])
                    except:
                        pass
                    
                    while True:
                        time.sleep(0.5)
                        try:
                            task = get_task(task_id)
                            if not task:
                                return True
                            
                            if task.get('interrupted'):
                                try:
                                    add_log(task_id, '任务已被中断', 'warning')
                                    update_task(task_id, status=TASK_STATUS['INTERRUPTED'])
                                except:
                                    pass
                                return True
                            
                            if not task.get('paused'):
                                try:
                                    add_log(task_id, '任务继续执行', 'info')
                                except:
                                    pass
                                return False
                        except:
                            # 暂停循环中出错，继续等待，不要让线程崩溃
                            time.sleep(1)
                            continue
            except Exception as e:
                # 检查中断/暂停时出错，不要让线程崩溃
                print(f"[警告] check_pause_interrupt出错: {e}")
                return False  # 出错时继续执行，不要中断整个任务
            
            return False
        
        items = task['items']
        amap_key = task.get('amap_key', '')
        doubao_api_key = task.get('doubao_api_key', '')
        doubao_model_id = task.get('doubao_model_id', '')
        task_provinces = task.get('provinces', [])  # 省份筛选列表
        
        add_log(task_id, f'开始处理任务，共 {len(items)} 条数据', 'info')
        if task_provinces:
            add_log(task_id, f'省份筛选已启用：{", ".join(task_provinces)}', 'info')
        
        # 第一阶段：搜索地址
        update_task(task_id, status=TASK_STATUS['SEARCHING'], current_phase='地址搜索')
        add_log(task_id, '=== 第一阶段：网络搜索地址 ===', 'info')
        
        for i, item in enumerate(items):
            if check_pause_interrupt():
                return
            
            if item.get('status') == ITEM_STATUS['COMPLETED']:
                continue
            
            item['status'] = ITEM_STATUS['SEARCHING']
            update_task(task_id, current_index=i + 1, current_item=item.get('name', ''))
            add_log(task_id, f'[{i+1}/{len(items)}] 正在搜索: {item.get("name", "")}', 'info')
            
            # 执行搜索（传入省份筛选参数和高德API Key）
            search_results = search_heritage_address(
                item.get('name', ''), 
                provinces=task_provinces,
                amap_key=amap_key,
                doubao_api_key=doubao_api_key,
                doubao_model_id=doubao_model_id
            )
            item['search_results'] = search_results
            
            # 提取推荐地址
            recommended_address = ''
            for r in search_results:
                if r.get('extracted_address'):
                    recommended_address = r['extracted_address']
                    break
            item['recommended_address'] = recommended_address
            item['status'] = ITEM_STATUS['WAITING_ADDRESS_REVIEW']
            
            add_log(task_id, f'  搜索完成，找到 {len(search_results)} 条结果', 'success')
            
            # 更新进度
            update_task(task_id, search_progress=int((i + 1) / len(items) * 100))
            
            # 全自动模式：使用多级搜索策略，根据置信度自动判断
            auto_mode = task.get('auto_mode', False)
            if auto_mode:
                # 使用多级搜索策略（优化版，传入省份筛选和高德API Key）
                multi_result = multi_level_search(
                    item.get('name', ''), 
                    min_confidence=0.7, 
                    provinces=task_provinces,
                    amap_key=amap_key,
                    doubao_api_key=doubao_api_key
                )
                
                # 保存搜索结果
                item['search_results'] = multi_result['results']
                item['search_level'] = multi_result['level']
                item['search_level_keyword'] = multi_result['level_keyword']
                item['best_address'] = multi_result['best_address']
                item['address_confidence'] = multi_result['best_confidence']
                item['result_status'] = multi_result['result_status']  # auto_approved/pending_review/manual_needed
                item['total_levels_tried'] = multi_result['total_levels_tried']
                item['place_name'] = multi_result['place_name']
                
                # 记录详细搜索日志
                add_log(task_id, f'  [多级搜索] 共尝试 {multi_result["total_levels_tried"]} 级关键词', 'info')
                for log_entry in multi_result['search_log']:
                    level_info = f'    第{log_entry["level"]}级: {log_entry["description"]}'
                    level_info += f' | 百度:{log_entry["baidu_result_count"]}条 必应:{log_entry["bing_result_count"]}条'
                    if log_entry.get('best_address'):
                        level_info += f' | 最佳地址: {log_entry["best_address"]}'
                    if log_entry.get('best_confidence', 0) > 0:
                        level_info += f' (置信度: {log_entry["best_confidence"]:.2f})'
                    if log_entry.get('fail_reason'):
                        level_info += f' | 失败原因: {log_entry["fail_reason"]}'
                    add_log(task_id, level_info, 'info')
                
                # 根据结果状态判断处理方式（优化版）
                result_status = multi_result['result_status']
                confidence = multi_result['best_confidence']
                best_addr = multi_result['best_address']
                
                if result_status == 'auto_approved' and best_addr:
                    # 高置信度（≥0.8）：自动确认
                    item['confirmed_address'] = best_addr
                    item['status'] = ITEM_STATUS['ADDRESS_APPROVED']
                    item['auto_approved'] = True
                    item['address_quality'] = 'high'
                    add_log(task_id, f'  [全自动] ✅ 自动确认地址: {best_addr} (置信度: {confidence:.2f}，第{multi_result["level"]}级)', 'success')
                elif result_status == 'pending_review' and best_addr:
                    # 中等置信度（0.5-0.8）：标记为待审核，建议人工复核
                    item['recommended_address'] = best_addr
                    item['status'] = ITEM_STATUS['WAITING_ADDRESS_REVIEW']
                    item['address_quality'] = 'medium'
                    add_log(task_id, f'  [全自动] ⚠️ 置信度中等，建议人工审核: {best_addr} (置信度: {confidence:.2f})', 'warning')
                    # 继续等待人工审核
                    add_log(task_id, f'  等待人工审核地址...', 'warning')
                    update_task(task_id, status=TASK_STATUS['WAITING_ADDRESS_REVIEW'])
                else:
                    # 低置信度（<0.5）或无结果：标记为需人工处理，跳过继续下一条
                    # 5级都失败才标记"需人工处理"
                    item['recommended_address'] = best_addr if best_addr else ''
                    item['status'] = ITEM_STATUS['WAITING_ADDRESS_REVIEW']
                    item['address_quality'] = 'low'
                    item['needs_manual'] = True
                    item['search_failed'] = True
                    
                    if multi_result['total_levels_tried'] >= 5:
                        add_log(task_id, f'  [全自动] ❌ 5级搜索全部失败，需人工处理 (置信度: {confidence:.2f})', 'error')
                    else:
                        add_log(task_id, f'  [全自动] ❌ 搜索失败，需人工处理 (置信度: {confidence:.2f})', 'error')
                    
                    # 为了不卡住流程，先标记为已确认但地址为空，后续坐标转换会失败
                    item['confirmed_address'] = best_addr if best_addr else ''
                    item['status'] = ITEM_STATUS['ADDRESS_APPROVED']
                    add_log(task_id, f'  [全自动] 跳过本条，继续处理下一条', 'warning')
                
                # 更新进度
                update_task(task_id, search_progress=int((i + 1) / len(items) * 100))
                
                # 如果是待审核状态，进入等待循环
                if item['status'] == ITEM_STATUS['WAITING_ADDRESS_REVIEW']:
                    # 等待审核完成（支持重新搜索）
                    while True:
                        time.sleep(0.5)
                        if check_pause_interrupt():
                            return
                        task = get_task(task_id)
                        item = task['items'][i]
                        
                        # 如果状态变回SEARCHING，说明用户点击了重新搜索
                        if item.get('status') == ITEM_STATUS['SEARCHING']:
                            add_log(task_id, f'  重新搜索中...', 'info')
                            # 重新执行多级搜索（支持自定义关键词，传入省份筛选）
                            custom_kw = item.get('custom_keyword', '')
                            search_keyword = custom_kw if custom_kw else item.get('name', '')
                            multi_result = multi_level_search(search_keyword, min_confidence=0.7, provinces=task_provinces, amap_key=amap_key, doubao_api_key=doubao_api_key, doubao_model_id=doubao_model_id)
                            with tasks_lock:
                                task['items'][i]['search_results'] = multi_result['results']
                                task['items'][i]['search_level'] = multi_result['level']
                                task['items'][i]['search_level_keyword'] = multi_result['level_keyword']
                                task['items'][i]['best_address'] = multi_result['best_address']
                                task['items'][i]['address_confidence'] = multi_result['best_confidence']
                                task['items'][i]['result_status'] = multi_result['result_status']
                                task['items'][i]['total_levels_tried'] = multi_result['total_levels_tried']
                                task['items'][i]['place_name'] = multi_result['place_name']
                                # 提取推荐地址
                                recommended_address = multi_result['best_address'] if multi_result['best_address'] else ''
                                task['items'][i]['recommended_address'] = recommended_address
                                # 设置地址质量
                                confidence = multi_result['best_confidence']
                                if confidence >= 0.7:
                                    task['items'][i]['address_quality'] = 'high'
                                elif confidence >= 0.5:
                                    task['items'][i]['address_quality'] = 'medium'
                                else:
                                    task['items'][i]['address_quality'] = 'low'
                                task['items'][i]['status'] = ITEM_STATUS['WAITING_ADDRESS_REVIEW']
                                # 清除自定义关键词
                                task['items'][i].pop('custom_keyword', None)
                            # 记录搜索日志
                            add_log(task_id, f'  [重新搜索] 共尝试 {multi_result["total_levels_tried"]} 级关键词', 'info')
                            for log_entry in multi_result['search_log']:
                                level_info = f'    第{log_entry["level"]}级: {log_entry["description"]}'
                                level_info += f' | 百度:{log_entry["baidu_result_count"]}条 必应:{log_entry["bing_result_count"]}条'
                                if log_entry.get('best_address'):
                                    level_info += f' | 最佳地址: {log_entry["best_address"]}'
                                if log_entry.get('best_confidence', 0) > 0:
                                    level_info += f' (置信度: {log_entry["best_confidence"]:.2f})'
                                add_log(task_id, level_info, 'info')
                            add_log(task_id, f'  重新搜索完成，找到 {len(multi_result["results"])} 条结果，最佳置信度: {confidence:.2f}', 'success')
                            continue
                        
                        if item.get('status') == ITEM_STATUS['ADDRESS_APPROVED']:
                            add_log(task_id, f'  地址已确认: {item.get("confirmed_address", "")}', 'success')
                            break
                    
                    update_task(task_id, status=TASK_STATUS['SEARCHING'])
                
                continue
            
            # 半自动模式：等待人工审核
            add_log(task_id, f'  等待人工审核地址...', 'warning')
            update_task(task_id, status=TASK_STATUS['WAITING_ADDRESS_REVIEW'])
            
            # 等待审核完成（支持重新搜索）
            while True:
                time.sleep(0.5)
                if check_pause_interrupt():
                    return
                task = get_task(task_id)
                item = task['items'][i]
                
                # 如果状态变回SEARCHING，说明用户点击了重新搜索
                if item.get('status') == ITEM_STATUS['SEARCHING']:
                    add_log(task_id, f'  重新搜索中...', 'info')
                    # 重新执行多级搜索（支持自定义关键词，传入省份筛选）
                    custom_kw = item.get('custom_keyword', '')
                    search_keyword = custom_kw if custom_kw else item.get('name', '')
                    multi_result = multi_level_search(search_keyword, min_confidence=0.7, provinces=task_provinces, amap_key=amap_key, doubao_api_key=doubao_api_key, doubao_model_id=doubao_model_id)
                    with tasks_lock:
                        task['items'][i]['search_results'] = multi_result['results']
                        task['items'][i]['search_level'] = multi_result['level']
                        task['items'][i]['search_level_keyword'] = multi_result['level_keyword']
                        task['items'][i]['best_address'] = multi_result['best_address']
                        task['items'][i]['address_confidence'] = multi_result['best_confidence']
                        task['items'][i]['result_status'] = multi_result['result_status']
                        task['items'][i]['total_levels_tried'] = multi_result['total_levels_tried']
                        task['items'][i]['place_name'] = multi_result['place_name']
                        # 提取推荐地址
                        recommended_address = multi_result['best_address'] if multi_result['best_address'] else ''
                        task['items'][i]['recommended_address'] = recommended_address
                        # 设置地址质量
                        confidence = multi_result['best_confidence']
                        if confidence >= 0.7:
                            task['items'][i]['address_quality'] = 'high'
                        elif confidence >= 0.5:
                            task['items'][i]['address_quality'] = 'medium'
                        else:
                            task['items'][i]['address_quality'] = 'low'
                        task['items'][i]['status'] = ITEM_STATUS['WAITING_ADDRESS_REVIEW']
                        # 清除自定义关键词
                        task['items'][i].pop('custom_keyword', None)
                    # 记录搜索日志
                    add_log(task_id, f'  [重新搜索] 共尝试 {multi_result["total_levels_tried"]} 级关键词', 'info')
                    for log_entry in multi_result['search_log']:
                        level_info = f'    第{log_entry["level"]}级: {log_entry["description"]}'
                        level_info += f' | 百度:{log_entry["baidu_result_count"]}条 必应:{log_entry["bing_result_count"]}条'
                        if log_entry.get('best_address'):
                            level_info += f' | 最佳地址: {log_entry["best_address"]}'
                        if log_entry.get('best_confidence', 0) > 0:
                            level_info += f' (置信度: {log_entry["best_confidence"]:.2f})'
                        add_log(task_id, level_info, 'info')
                    add_log(task_id, f'  重新搜索完成，找到 {len(multi_result["results"])} 条结果，最佳置信度: {confidence:.2f}', 'success')
                    continue
                
                if item.get('status') == ITEM_STATUS['ADDRESS_APPROVED']:
                    add_log(task_id, f'  地址已确认: {item.get("confirmed_address", "")}', 'success')
                    break
            
            update_task(task_id, status=TASK_STATUS['SEARCHING'])
        
        # 两遍模式第一遍完成：不继续坐标转换，等待用户触发第二遍
        if 'two_pass_pass1' in dir() and two_pass_pass1:
            with tasks_lock:
                task = get_task(task_id)
                if task:
                    task['pass'] = 2
                    task['status'] = 'pass1_done'
                    task['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            add_log(task_id, '=== 第一遍搜索完成，等待第二遍审核 ===', 'success')
            return  # 不要继续坐标转换
        # 第二阶段：坐标转换
        update_task(task_id, status=TASK_STATUS['GEOCODING'], current_phase='坐标转换')
        add_log(task_id, '=== 第二阶段：高德地图坐标转换 ===', 'info')
        
        for i, item in enumerate(items):
            if check_pause_interrupt():
                return
            
            if item.get('status') == ITEM_STATUS['COMPLETED']:
                continue
            
            item['status'] = ITEM_STATUS['GEOCODING']
            update_task(task_id, current_index=i + 1, current_item=item.get('name', ''))
            add_log(task_id, f'[{i+1}/{len(items)}] 正在转换坐标: {item.get("confirmed_address", "")}', 'info')
            
            # 执行坐标转换（传入省份筛选）
            geo_results = geocode_address(item.get('confirmed_address', ''), amap_key, provinces=task_provinces)
            item['geo_results'] = geo_results
            
            # 推荐第一个结果
            if geo_results and geo_results[0].get('longitude', 0) > 0:
                item['recommended_coord'] = {
                    'longitude': geo_results[0]['longitude'],
                    'latitude': geo_results[0]['latitude'],
                    'formatted_address': geo_results[0]['formatted_address']
                }
            
            item['status'] = ITEM_STATUS['WAITING_COORD_REVIEW']
            
            add_log(task_id, f'  坐标转换完成，找到 {len(geo_results)} 条结果', 'success')
            
            # 更新进度
            update_task(task_id, geo_progress=int((i + 1) / len(items) * 100))
            
            # 全自动模式：自动选择第一个结果
            auto_mode = task.get('auto_mode', False)
            if auto_mode:
                if geo_results and geo_results[0].get('longitude', 0) > 0:
                    item['confirmed_longitude'] = geo_results[0]['longitude']
                    item['confirmed_latitude'] = geo_results[0]['latitude']
                    item['status'] = ITEM_STATUS['COMPLETED']
                    add_log(task_id, f'  [全自动] 自动选择坐标: {geo_results[0]["longitude"]}, {geo_results[0]["latitude"]}', 'success')
                else:
                    item['confirmed_longitude'] = 0
                    item['confirmed_latitude'] = 0
                    item['status'] = ITEM_STATUS['COMPLETED']
                    add_log(task_id, f'  [全自动] 坐标转换失败，标记为0', 'warning')
                continue
            
            # 半自动模式：等待人工审核
            add_log(task_id, f'  等待人工审核坐标...', 'warning')
            update_task(task_id, status=TASK_STATUS['WAITING_COORD_REVIEW'])
            
            # 等待审核完成（支持重新转换）
            while True:
                time.sleep(0.5)
                if check_pause_interrupt():
                    return
                task = get_task(task_id)
                item = task['items'][i]
                
                # 如果状态变回GEOCODING，说明用户点击了重新转换
                if item.get('status') == ITEM_STATUS['GEOCODING']:
                    add_log(task_id, f'  重新转换坐标中...', 'info')
                    # 重新执行坐标转换（传入省份筛选）
                    geo_results = geocode_address(item.get('confirmed_address', ''), amap_key, provinces=task_provinces)
                    with tasks_lock:
                        task['items'][i]['geo_results'] = geo_results
                        # 推荐第一个结果
                        if geo_results and geo_results[0].get('longitude', 0) > 0:
                            task['items'][i]['recommended_coord'] = {
                                'longitude': geo_results[0]['longitude'],
                                'latitude': geo_results[0]['latitude'],
                                'formatted_address': geo_results[0]['formatted_address']
                            }
                        task['items'][i]['status'] = ITEM_STATUS['WAITING_COORD_REVIEW']
                    add_log(task_id, f'  重新转换完成，找到 {len(geo_results)} 条结果', 'success')
                    continue
                
                if item.get('status') == ITEM_STATUS['COORD_APPROVED']:
                    add_log(task_id, f'  坐标已确认: {item.get("confirmed_longitude")}, {item.get("confirmed_latitude")}', 'success')
                    item['status'] = ITEM_STATUS['COMPLETED']
                    break
            
            update_task(task_id, status=TASK_STATUS['GEOCODING'])
        
        # 全部完成
        update_task(task_id, status=TASK_STATUS['COMPLETED'], 
                    search_progress=100, geo_progress=100,
                    current_phase='已完成')
        add_log(task_id, '=== 所有数据处理完成 ===', 'success')
        
    except Exception as e:
        # 确保任务出错时不会影响整个服务
        try:
            add_log(task_id, f'任务执行出错: {str(e)}', 'error')
            update_task(task_id, status=TASK_STATUS['ERROR'], error=str(e))
        except:
            # 如果连日志记录都失败了，至少打印到控制台
            print(f"[严重错误] 任务 {task_id} 执行失败: {e}")
            import traceback
            traceback.print_exc()


# ==================== 路由定义 ====================

@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传Excel文件"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': '没有选择文件'}), 400
    
    if file:
        filename = secure_filename(file.filename)
        task_id = generate_task_id()
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_{filename}')
        file.save(filepath)
        
        try:
            # 读取Excel获取列名
            df = pd.read_excel(filepath, engine='openpyxl')
            columns = df.columns.tolist()
            row_count = len(df)
            
            # 创建任务
            with tasks_lock:
                tasks[task_id] = {
                    'task_id': task_id,
                    'filename': filename,
                    'filepath': filepath,
                    'columns': columns,
                    'row_count': row_count,
                    'status': TASK_STATUS['UPLOADED'],
                    'items': [],
                    'field_mapping': {},
                    'amap_key': '',
                    'doubao_api_key': '',
                    'doubao_model_id': '',
                    'auto_mode': False,  # 处理模式：False=半自动，True=全自动
                    'mode': 'step_by_step',  # 处理模式：step_by_step=逐条审核，two_pass=两遍模式
                    'pass': 1,                # 当前遍数：1=第一遍，2=第二遍
                    'provinces': [],     # 省份筛选（空=全国）
                    'logs': [],
                    'current_index': 0,
                    'current_item': '',
                    'current_phase': '',
                    'search_progress': 0,
                    'geo_progress': 0,
                    'paused': False,
                    'interrupted': False,
                    'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
            
            add_log(task_id, f'文件上传成功: {filename}，共 {row_count} 行数据', 'success')
            add_log(task_id, f'检测到 {len(columns)} 个字段: {", ".join(columns)}', 'info')
            
            return jsonify({
                'success': True,
                'task_id': task_id,
                'columns': columns,
                'row_count': row_count,
                'filename': filename
            })
            
        except Exception as e:
            return jsonify({'success': False, 'message': f'读取Excel失败: {str(e)}'}), 400


@app.route('/api/field-mapping', methods=['POST'])
def set_field_mapping():
    """设置字段映射"""
    data = request.json
    task_id = data.get('task_id')
    field_mapping = data.get('field_mapping', {})
    amap_key = data.get('amap_key', '')
    doubao_api_key = data.get('doubao_api_key', '')
    doubao_model_id = data.get('doubao_model_id', '').strip()
    auto_mode = data.get('auto_mode', False)  # 处理模式
    provinces = data.get('provinces', [])      # 省份筛选列表（空=全国）
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    try:
        # 读取Excel数据
        df = pd.read_excel(task['filepath'], engine='openpyxl')
        
        # 根据映射提取数据
        items = []
        name_field = field_mapping.get('name', '')
        
        for idx, row in df.iterrows():
            item = {
                'id': idx,
                'name': str(row.get(name_field, '')) if name_field else '',
                'original_data': row.to_dict(),
                'status': ITEM_STATUS['PENDING'],
                'search_results': [],
                'recommended_address': '',
                'confirmed_address': '',
                'geo_results': [],
                'recommended_coord': None,
                'confirmed_longitude': None,
                'confirmed_latitude': None,
                'error': ''
            }
            items.append(item)
        
        # 更新任务
        update_task(task_id, 
                    field_mapping=field_mapping,
                    amap_key=amap_key,
                    doubao_api_key=doubao_api_key,
                    doubao_model_id=doubao_model_id,
                    auto_mode=auto_mode,
                    provinces=provinces,
                    items=items,
                    status=TASK_STATUS['MAPPED'])
        
        add_log(task_id, f'字段映射完成，名称字段: {name_field}', 'success')
        add_log(task_id, f'高德API Key已配置: {"是" if amap_key else "否"}', 'info')
        if provinces:
            add_log(task_id, f'搜索范围限定省份: {", ".join(provinces)}', 'info')
        else:
            add_log(task_id, '搜索范围：全国（未限定省份）', 'info')
        
        return jsonify({
            'success': True,
            'message': '字段映射成功',
            'item_count': len(items)
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'字段映射失败: {str(e)}'}), 400


@app.route('/api/start', methods=['POST'])
def start_task():
    """开始任务"""
    data = request.json
    task_id = data.get('task_id')
    mode = data.get('mode', 'step_by_step')  # 处理模式：step_by_step=逐条审核，two_pass=两遍模式
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    # 重置状态
    update_task(task_id, paused=False, interrupted=False, mode=mode)
    with tasks_lock:
        t = get_task(task_id)
        if t:
            t['pass'] = 1
            t['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 启动后台线程（根据模式选择函数）
    if mode == 'two_pass':
        thread = threading.Thread(target=process_task_two_pass, args=(task_id,), daemon=True)
    else:
        thread = threading.Thread(target=process_task, args=(task_id,), daemon=True)
    thread.start()
    
    add_log(task_id, f'任务已启动，模式：{"两遍模式" if mode == "two_pass" else "逐条审核"}', 'info')
    
    return jsonify({'success': True, 'message': '任务已启动', 'mode': mode})


@app.route('/api/start_pass2', methods=['POST'])
def start_pass2():
    """开始第二遍审核（两遍模式专用）"""
    data = request.json
    task_id = data.get('task_id')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    mode = task.get('mode', 'step_by_step')
    if mode != 'two_pass':
        return jsonify({'success': False, 'message': '当前任务不是两遍模式'}), 400
    
    if task.get('status') != 'pass1_done':
        return jsonify({'success': False, 'message': f'第一遍尚未完成（当前状态：{task.get("status")}）'}), 400
    
    # 设置第二遍状态
    with tasks_lock:
        t = get_task(task_id)
        if t:
            t['pass'] = 2
            t['status'] = 'pass2_reviewing'
            t['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    add_log(task_id, '=== 开始第二遍审核 ===', 'success')
    
    return jsonify({'success': True, 'message': '第二遍审核已启动'})


@app.route('/api/pause', methods=['POST'])
def pause_task():
    """暂停任务"""
    data = request.json
    task_id = data.get('task_id')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    update_task(task_id, paused=True)
    add_log(task_id, '收到暂停指令', 'warning')
    
    return jsonify({'success': True, 'message': '暂停指令已发送'})


@app.route('/api/resume', methods=['POST'])
def resume_task():
    """继续任务"""
    data = request.json
    task_id = data.get('task_id')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    update_task(task_id, paused=False)
    add_log(task_id, '收到继续指令', 'info')
    
    return jsonify({'success': True, 'message': '继续指令已发送'})


@app.route('/api/interrupt', methods=['POST'])
def interrupt_task():
    """中断任务"""
    data = request.json
    task_id = data.get('task_id')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    update_task(task_id, interrupted=True, paused=False)
    add_log(task_id, '收到中断指令', 'error')
    
    return jsonify({'success': True, 'message': '中断指令已发送'})

@app.route('/api/return-to-config', methods=['POST'])
def return_to_config():
    """返回配置页面，支持保存进度或丢弃进度"""
    data = request.json
    task_id = data.get('task_id')
    save_progress = data.get('save_progress', False)
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    # 先中断当前正在运行的任务
    update_task(task_id, interrupted=True, paused=False)
    add_log(task_id, '返回配置页面，停止当前任务', 'warning')
    
    if save_progress:
        # 保存进度模式：保留已处理的items数据
        add_log(task_id, '已保存当前处理进度', 'success')
        # 状态改为已映射，保留所有数据
        update_task(task_id, status=TASK_STATUS['MAPPED'], 
                    paused=False, interrupted=False)
    else:
        # 丢弃进度模式：清空items数据，保留字段映射和API Key
        add_log(task_id, '已丢弃当前处理进度，重新开始', 'warning')
        # 重新读取Excel数据，重置items
        try:
            df = pd.read_excel(task['filepath'], engine='openpyxl')
            field_mapping = task.get('field_mapping', {})
            name_field = field_mapping.get('name', '')
            
            items = []
            for idx, row in df.iterrows():
                item = {
                    'id': idx,
                    'name': str(row.get(name_field, '')) if name_field else '',
                    'original_data': row.to_dict(),
                    'status': ITEM_STATUS['PENDING'],
                    'search_results': [],
                    'recommended_address': '',
                    'confirmed_address': '',
                    'geo_results': [],
                    'recommended_coord': None,
                    'confirmed_longitude': None,
                    'confirmed_latitude': None,
                    'error': ''
                }
                items.append(item)
            
            update_task(task_id, 
                        items=items,
                        status=TASK_STATUS['MAPPED'],
                        paused=False,
                        interrupted=False,
                        search_progress=0,
                        geo_progress=0,
                        current_index=0,
                        current_item='',
                        current_phase='')
        except Exception as e:
            add_log(task_id, f'重置数据失败: {str(e)}', 'error')
            return jsonify({'success': False, 'message': f'重置数据失败: {str(e)}'}), 400
    
    return jsonify({
        'success': True, 
        'message': '已返回配置页面',
        'field_mapping': task.get('field_mapping', {}),
        'amap_key': task.get('amap_key', ''),
        'auto_mode': task.get('auto_mode', False),
        'provinces': task.get('provinces', [])
    })


@app.route('/api/retry-item', methods=['POST'])
def retry_item():
    """单条重试"""
    data = request.json
    task_id = data.get('task_id')
    item_id = data.get('item_id')
    retry_type = data.get('type', 'search')  # search 或 geo
    custom_keyword = data.get('custom_keyword', '')  # 自定义搜索关键词
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    with tasks_lock:
        if item_id < len(task['items']):
            item = task['items'][item_id]
            if retry_type == 'search':
                item['status'] = ITEM_STATUS['SEARCHING']
                item['search_results'] = []
                item['recommended_address'] = ''
                item['custom_keyword'] = custom_keyword  # 保存自定义关键词
                if custom_keyword:
                    add_log(task_id, f'使用自定义关键词重新搜索第 {item_id + 1} 条: {custom_keyword}', 'info')
                else:
                    add_log(task_id, f'重新搜索第 {item_id + 1} 条: {item.get("name", "")}', 'info')
            else:
                item['status'] = ITEM_STATUS['GEOCODING']
                item['geo_results'] = []
                item['recommended_coord'] = None
                add_log(task_id, f'重新转换坐标第 {item_id + 1} 条', 'info')
    
    return jsonify({'success': True, 'message': '重试指令已发送'})


@app.route('/api/review-address', methods=['POST'])
def review_address():
    """审核地址"""
    data = request.json
    task_id = data.get('task_id')
    item_id = data.get('item_id')
    confirmed_address = data.get('confirmed_address', '')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    with tasks_lock:
        if item_id < len(task['items']):
            item = task['items'][item_id]
            item['confirmed_address'] = confirmed_address
            item['status'] = ITEM_STATUS['ADDRESS_APPROVED']
    
    add_log(task_id, f'第 {item_id + 1} 条地址审核通过', 'success')
    
    return jsonify({'success': True, 'message': '地址已确认'})


@app.route('/api/review-coord', methods=['POST'])
def review_coord():
    """审核坐标"""
    data = request.json
    task_id = data.get('task_id')
    item_id = data.get('item_id')
    longitude = data.get('longitude')
    latitude = data.get('latitude')
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    with tasks_lock:
        if item_id < len(task['items']):
            item = task['items'][item_id]
            item['confirmed_longitude'] = longitude
            item['confirmed_latitude'] = latitude
            item['status'] = ITEM_STATUS['COORD_APPROVED']
    
    add_log(task_id, f'第 {item_id + 1} 条坐标审核通过', 'success')
    
    return jsonify({'success': True, 'message': '坐标已确认'})


@app.route('/api/status/<task_id>')
def get_status(task_id):
    """获取任务状态"""
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    # 返回任务状态（不返回原始数据，避免数据过大）
    status_info = {
        'task_id': task['task_id'],
        'status': task['status'],
        'current_index': task['current_index'],
        'current_item': task['current_item'],
        'current_phase': task['current_phase'],
        'search_progress': task['search_progress'],
        'geo_progress': task['geo_progress'],
        'total_items': len(task['items']),
        'logs': task['logs'][-100:],  # 只返回最近100条
        'updated_at': task['updated_at'],
        'paused': task.get('paused', False),
        'interrupted': task.get('interrupted', False),
        'auto_mode': task.get('auto_mode', False)
    }
    
    # 检查是否有待审核的项目
    waiting_address_items = []
    waiting_coord_items = []
    for i, item in enumerate(task['items']):
        if item['status'] == ITEM_STATUS['WAITING_ADDRESS_REVIEW']:
            waiting_address_items.append({
                'id': i,
                'name': item['name'],
                'search_results': item['search_results'],
                'recommended_address': item['recommended_address']
            })
        elif item['status'] == ITEM_STATUS['WAITING_COORD_REVIEW']:
            waiting_coord_items.append({
                'id': i,
                'name': item['name'],
                'confirmed_address': item['confirmed_address'],
                'geo_results': item['geo_results'],
                'recommended_coord': item['recommended_coord']
            })
    
    status_info['waiting_address_items'] = waiting_address_items
    status_info['waiting_coord_items'] = waiting_coord_items
    
    return jsonify({'success': True, 'data': status_info})


@app.route('/api/items/<task_id>')
def get_items(task_id):
    """获取所有项目列表"""
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    items_summary = []
    for i, item in enumerate(task['items']):
        items_summary.append({
            'id': i,
            'name': item['name'],
            'status': item['status'],
            'confirmed_address': item.get('confirmed_address', ''),
            'longitude': item.get('confirmed_longitude'),
            'latitude': item.get('confirmed_latitude'),
            'address_quality': item.get('address_quality', ''),  # high/medium/low
            'address_confidence': item.get('address_confidence', 0),
            'auto_approved': item.get('auto_approved', False),
            'search_failed': item.get('search_failed', False),
            'needs_manual': item.get('needs_manual', False)
        })
    
    return jsonify({'success': True, 'items': items_summary})


@app.route('/api/export/<task_id>')
def export_excel(task_id):
    """导出Excel"""
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    try:
        # 构建导出数据
        export_data = []
        for item in task['items']:
            row = item.get('original_data', {}).copy()
            row['确认地址'] = item.get('confirmed_address', '')
            row['经度'] = item.get('confirmed_longitude', '')
            row['纬度'] = item.get('confirmed_latitude', '')
            row['处理状态'] = item.get('status', '')
            export_data.append(row)
        
        df = pd.DataFrame(export_data)
        
        # 保存到临时文件
        export_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_result.xlsx')
        df.to_excel(export_path, index=False)
        
        add_log(task_id, f'结果已导出，共 {len(export_data)} 条数据', 'success')
        
        return send_file(export_path, 
                         as_attachment=True,
                         download_name=f'非遗地址坐标结果_{task_id}.xlsx')
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'导出失败: {str(e)}'}), 400


@app.route('/api/tasks')
def list_tasks():
    """获取所有任务列表"""
    with tasks_lock:
        task_list = []
        for task_id, task in tasks.items():
            task_list.append({
                'task_id': task_id,
                'filename': task['filename'],
                'status': task['status'],
                'total_items': len(task['items']),
                'created_at': task['created_at']
            })
        # 按创建时间倒序
        task_list.sort(key=lambda x: x['created_at'], reverse=True)
    
    return jsonify({'success': True, 'tasks': task_list})




@app.route("/api/search_with_ai", methods=["POST"])
def search_with_ai():
    """
    使用豆包AI搜索非遗项目地址
    请求体: {project_name, doubao_api_key, amap_key}
    """
    try:
        data = request.get_json()
        project_name = data.get("project_name", "").strip()
        doubao_api_key = data.get("doubao_api_key", "").strip()
        doubao_model_id = data.get("doubao_model_id", "").strip()
        amap_key = data.get("amap_key", "").strip() or None

        if not project_name:
            return jsonify({"success": False, "message": "缺少项目名称"}), 400
        if not doubao_api_key:
            return jsonify({"success": False, "message": "缺少豆包API Key，请在配置面板填写"}), 400
        if not doubao_model_id:
            return jsonify({"success": False, "message": "缺少推理接入点ID，请在配置面板填写（格式：ep-xxxxxxxxxx-xxxxx）"}), 400

        # 调用豆包AI
        try:
            result = call_doubao_for_address(project_name, doubao_api_key, amap_key, model_id=doubao_model_id)
        except Exception as api_err:
            print(f"[search_with_ai] API调用异常: {api_err}")
            return jsonify({"success": False, "message": f"豆包API调用失败: {str(api_err)}"}), 500

        if result:
            return jsonify({
                "success": True,
                "address": result["address"],
                "lng": result.get("lng"),
                "lat": result.get("lat"),
                "confidence": result["confidence"],
                "source": result["source"],
                "raw_response": result.get("raw_response", "")
            })
        else:
            return jsonify({
                "success": False,
                "message": "AI返回了结果但无法解析出有效地址，请检查模型ID是否正确"
            }), 404

    except Exception as e:
        return jsonify({"success": False, "message": f"AI搜索失败: {str(e)}"}), 500



# ========== 两遍模式专用函数 ==========

def process_task_two_pass(task_id):
    """两遍模式处理线程（由 /api/start 根据 mode 参数调用）"""
    try:
        task = get_task(task_id)
        if not task:
            return
        pass_num = task.get('pass', 1)
        if pass_num == 1:
            _two_pass_search_all(task_id, task)
        else:
            _two_pass_review_mode(task_id, task)
    except Exception as e:
        import traceback
        add_log(task_id, f'两遍模式处理出错：{str(e)}', 'error')
        traceback.print_exc()

def _two_pass_search_all(task_id, task):
    """第一遍：批量自动搜索（不阻塞等待审核）"""
    items = task['items']
    amap_key = task.get('amap_key', '')
    doubao_api_key = task.get('doubao_api_key', '')
    doubao_model_id = task.get('doubao_model_id', '')
    add_log(task_id, f'两遍模式第一遍：开始批量自动搜索（共 {len(items)} 条）', 'info')
    for idx, item in enumerate(items):
        if item.get('status') == '已完成':
            continue
        if task.get('paused'):
            while task.get('paused') and not task.get('interrupted'):
                time.sleep(0.5)
            if task.get('interrupted'):
                add_log(task_id, '任务被中断', 'warning')
                return
        item['status'] = '搜索中'
        item['pass'] = 1
        item['search_count'] = 1
        keyword = item.get('project_name', '')
        results = search_heritage_address(
            keyword,
            provinces=item.get('province'),
            amap_key=amap_key,
            doubao_api_key=doubao_api_key,
            doubao_model_id=doubao_model_id
        )
        item['search_results'] = results
        item['search_history'] = [{
            'pass': 1,
            'count': 1,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'results': results
        }]
        best = results[0] if results else None
        if best and best.get('confidence', 0) >= 0.8:
            item['confirmed_address'] = best.get('address')
            item['confirmed_lng'] = best.get('lng')
            item['confirmed_lat'] = best.get('lat')
            item['status'] = '第一遍自动确认'
            item['auto_approved'] = True
            add_log(task_id, f'  ✅ 自动确认：{best.get("address")} (置信度: {best.get("confidence", 0):.2f})', 'success')
        else:
            item['status'] = '等待地址审核'
            item['auto_approved'] = False
            add_log(task_id, f'  需要人工审核（找到 {len(results)} 条结果）', 'warning')
        task['progress'] = int((idx + 1) / len(items) * 50)
        add_log(task_id, f'第一遍：已处理 {idx + 1}/{len(items)}', 'info')
    task['pass'] = 1
    task['status'] = 'pass1_done'
    task['progress'] = 50
    add_log(task_id, '=== 第一遍搜索完成，等待第二遍审核 ===', 'success')

def _two_pass_review_mode(task_id, task):
    """第二遍：人工审核模式（设置状态，由前端驱动审核）"""
    add_log(task_id, '两遍模式第二遍：进入人工审核模式', 'info')
    task['status'] = 'pass2_reviewing'


if __name__ == '__main__':
    print('=' * 60)
    print('非遗地址查找工具 - 启动中...')
    print('访问地址: http://localhost:5000')
    print('=' * 60)
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)

