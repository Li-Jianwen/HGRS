#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HGR Discover — API 发现脚本
直接调用人遗系统后端 API 获取全部批次文件信息，无需浏览器。

用法:
    python hgr_discover.py --output-dir <输出目录>
    或
    python hgr_discover.py --output-dir <输出目录> --cookie <cookie_string>

输出:
    items.json    — [{title, pdf_url, publicity_file_id}]
    cookies.json  — 占位文件（后续下载需手动补充 cookie）
"""
import sys, io, os, json, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import requests

API_URL = 'https://apply.hgrg.net/api/backend/projectPublicity/fileInfo'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Content-Type': 'application/json',
    'Referer': 'https://apply.hgrg.net/login',
}

# 过滤规则：文件名包含"批审批结果"，排除"审查结果"/"补充公示"/"计划通知"/"受理截止"/"会议时间"
EXCLUDE_KEYWORDS = ['审查结果', '补充公示', '计划通知', '受理截止', '会议时间']

def discover(output_dir: str, cookie_str: str = ''):
    print(f"🌐 调用 API: {API_URL}")
    
    cookies = {}
    if cookie_str:
        for part in cookie_str.split(';'):
            if '=' in part:
                k, v = part.strip().split('=', 1)
                cookies[k] = v
    
    r = requests.post(API_URL, json={}, headers=HEADERS, cookies=cookies, timeout=30)
    r.encoding = 'utf-8'
    
    if r.status_code != 200:
        print(f"❌ API 返回 {r.status_code}: {r.text[:200]}")
        print("   如需 Cookie，请登录后从浏览器复制并传入 --cookie")
        return
    
    raw = r.json()
    # API 返回格式: {"message": "...", "code": "...", "data": [...]}
    if isinstance(raw, dict) and 'data' in raw:
        records = raw['data']
    else:
        records = raw
    print(f"📋 API 返回 {len(records)} 条记录")
    
    items = []
    for rec in records:
        filename = rec.get('fileName', '')
        # 过滤条件
        if '批审批结果' not in filename:
            continue
        if any(kw in filename for kw in EXCLUDE_KEYWORDS):
            continue
        
        pdf_url = rec.get('publicityFileUrl', '')
        if pdf_url and not pdf_url.startswith('http'):
            pdf_url = 'https://apply.hgrg.net' + pdf_url
        
        items.append({
            'title': filename,
            'pdf_url': pdf_url,
            'publicity_file_id': rec.get('publicityFileId', ''),
        })
    
    print(f"✅ 过滤后: {len(items)} 个审批结果批次")
    for i, item in enumerate(items[:5]):
        print(f"   [{i+1}] {item['title'][:60]}...")
    if len(items) > 5:
        print(f"   ... 共 {len(items)} 条")
    
    # 保存 items.json
    items_path = os.path.join(output_dir, 'items.json')
    with open(items_path, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    print(f"📝 已保存: {items_path}")
    
    # 保存 cookies.json（占位）
    cookies_path = os.path.join(output_dir, 'cookies.json')
    with open(cookies_path, 'w', encoding='utf-8') as f:
        json.dump(cookies, f, ensure_ascii=False, indent=2)
    print(f"📝 已保存: {cookies_path} (Cookie {'有' if cookies else '空，需手动补充'})")

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', required=True, help='输出目录')
    parser.add_argument('--cookie', default='', help='Cookie 字符串（可选）')
    args = parser.parse_args()
    discover(args.output_dir, args.cookie)
