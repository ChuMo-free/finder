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
    'ERROR': '出错'
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
def search_heritage_address(keyword):
    """
    搜索非遗项目的真实地址
    使用百度搜索结果，提取地址信息
    """
    results = []
    try:
        # 构造搜索关键词
        search_query = f"{keyword} 地址 非遗"
        
        # 使用百度搜索
        url = f"https://www.baidu.com/s?wd={requests.utils.quote(search_query)}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'
        
        soup = BeautifulSoup(response.text, 'lxml')
        
        # 提取搜索结果
        search_results = soup.select('.result')[:5]  # 取前5条结果
        
        for idx, result in enumerate(search_results):
            title_elem = result.select_one('h3 a')
            snippet_elem = result.select_one('.c-abstract')
            
            title = title_elem.get_text(strip=True) if title_elem else ''
            snippet = snippet_elem.get_text(strip=True) if snippet_elem else ''
            link = title_elem.get('href', '') if title_elem else ''
            
            # 尝试从摘要中提取地址
            address = extract_address_from_text(title + ' ' + snippet)
            
            results.append({
                'index': idx + 1,
                'title': title,
                'snippet': snippet,
                'link': link,
                'extracted_address': address,
                'source': '百度搜索'
            })
        
        # 如果没找到结果，尝试第二种搜索方式
        if not results:
            results = fallback_search(keyword)
            
    except Exception as e:
        results.append({
            'index': 1,
            'title': f'搜索出错: {str(e)}',
            'snippet': '请手动输入地址',
            'link': '',
            'extracted_address': '',
            'source': '错误'
        })
    
    return results


def fallback_search(keyword):
    """备用搜索方式"""
    results = []
    try:
        # 尝试使用必应搜索
        search_query = f"{keyword} 非遗 地址"
        url = f"https://cn.bing.com/search?q={requests.utils.quote(search_query)}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'lxml')
        
        items = soup.select('.b_algo')[:5]
        for idx, item in enumerate(items):
            title_elem = item.select_one('h2 a')
            snippet_elem = item.select_one('.b_caption p')
            
            title = title_elem.get_text(strip=True) if title_elem else ''
            snippet = snippet_elem.get_text(strip=True) if snippet_elem else ''
            link = title_elem.get('href', '') if title_elem else ''
            
            address = extract_address_from_text(title + ' ' + snippet)
            
            results.append({
                'index': idx + 1,
                'title': title,
                'snippet': snippet,
                'link': link,
                'extracted_address': address,
                'source': '必应搜索'
            })
    except Exception:
        pass
    
    return results


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


# ==================== 高德地图API功能 ====================
def geocode_address(address, amap_key):
    """
    使用高德地图API将地址转换为坐标
    返回坐标结果列表
    """
    results = []
    try:
        url = 'https://restapi.amap.com/v3/geocode/geo'
        params = {
            'key': amap_key,
            'address': address,
            'output': 'json'
        }
        
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
    后台任务处理线程
    按照工作流逐步处理数据
    """
    task = get_task(task_id)
    if not task:
        return
    
    try:
        # 检查是否需要暂停或中断
        def check_pause_interrupt():
            task = get_task(task_id)
            if task.get('interrupted'):
                add_log(task_id, '任务已被中断', 'warning')
                update_task(task_id, status=TASK_STATUS['INTERRUPTED'])
                return True
            if task.get('paused'):
                add_log(task_id, '任务已暂停', 'info')
                update_task(task_id, status=TASK_STATUS['PAUSED'])
                while True:
                    time.sleep(0.5)
                    task = get_task(task_id)
                    if task.get('interrupted'):
                        add_log(task_id, '任务已被中断', 'warning')
                        update_task(task_id, status=TASK_STATUS['INTERRUPTED'])
                        return True
                    if not task.get('paused'):
                        add_log(task_id, '任务继续执行', 'info')
                        return False
            return False
        
        items = task['items']
        amap_key = task.get('amap_key', '')
        
        add_log(task_id, f'开始处理任务，共 {len(items)} 条数据', 'info')
        
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
            
            # 执行搜索
            search_results = search_heritage_address(item.get('name', ''))
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
            
            # 等待人工审核
            add_log(task_id, f'  等待人工审核地址...', 'warning')
            update_task(task_id, status=TASK_STATUS['WAITING_ADDRESS_REVIEW'])
            
            # 等待审核完成
            while True:
                time.sleep(0.5)
                if check_pause_interrupt():
                    return
                task = get_task(task_id)
                item = task['items'][i]
                if item.get('status') == ITEM_STATUS['ADDRESS_APPROVED']:
                    add_log(task_id, f'  地址已确认: {item.get("confirmed_address", "")}', 'success')
                    break
            
            update_task(task_id, status=TASK_STATUS['SEARCHING'])
        
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
            
            # 执行坐标转换
            geo_results = geocode_address(item.get('confirmed_address', ''), amap_key)
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
            
            # 等待人工审核
            add_log(task_id, f'  等待人工审核坐标...', 'warning')
            update_task(task_id, status=TASK_STATUS['WAITING_COORD_REVIEW'])
            
            # 等待审核完成
            while True:
                time.sleep(0.5)
                if check_pause_interrupt():
                    return
                task = get_task(task_id)
                item = task['items'][i]
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
        add_log(task_id, f'任务执行出错: {str(e)}', 'error')
        update_task(task_id, status=TASK_STATUS['ERROR'], error=str(e))


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
            df = pd.read_excel(filepath)
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
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    try:
        # 读取Excel数据
        df = pd.read_excel(task['filepath'])
        
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
                    items=items,
                    status=TASK_STATUS['MAPPED'])
        
        add_log(task_id, f'字段映射完成，名称字段: {name_field}', 'success')
        add_log(task_id, f'高德API Key已配置: {"是" if amap_key else "否"}', 'info')
        
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
    
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    
    # 重置状态
    update_task(task_id, paused=False, interrupted=False)
    
    # 启动后台线程
    thread = threading.Thread(target=process_task, args=(task_id,), daemon=True)
    thread.start()
    
    add_log(task_id, '任务已启动', 'info')
    
    return jsonify({'success': True, 'message': '任务已启动'})


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
            df = pd.read_excel(task['filepath'])
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
        'amap_key': task.get('amap_key', '')
    })


@app.route('/api/retry-item', methods=['POST'])
def retry_item():
    """单条重试"""
    data = request.json
    task_id = data.get('task_id')
    item_id = data.get('item_id')
    retry_type = data.get('type', 'search')  # search 或 geo
    
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
        'interrupted': task.get('interrupted', False)
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
            'latitude': item.get('confirmed_latitude')
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


if __name__ == '__main__':
    print('=' * 60)
    print('非遗地址查找工具 - 启动中...')
    print('访问地址: http://localhost:5000')
    print('=' * 60)
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
