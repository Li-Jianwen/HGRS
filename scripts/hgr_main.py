#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HGR Main
主控脚本：读取配置 → 增量判断 → 处理新批次 → 输出简报
"""

import os
import sys
import io
import json
import logging
import configparser
from typing import List, Dict, Tuple, Optional, Any
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from hgr_processor import HGRProcessor
from excel_writer import HGRWriter

# 默认配置
DEFAULT_CONFIG = {
    'base_url': 'https://apply.hgrg.net/login',
    'data_dir': './data',
    'summary_filename': '汇总_中国人类遗传资源行政许可事项.xlsx',
    'batch_filename_template': '中国人类遗传资源行政许可事项{year}年第{batch}批审批结果公示.xlsx',
    'download_retry': '3',
    'download_timeout': '30',
    'browser_timeout': '30',
    'log_file': 'hgr.log',
}

class HGRMain:
    def __init__(self, config_path: str = '../config.ini'):
        self.config_path = config_path
        
        # Setup logging
        self.logger = logging.getLogger('hgr-main')
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False  # 防止日志传播到根 logger 导致重复
        
        self.config = self.load_config()
        self.data_dir_full = self.get_full_path(self.config.get('data_dir'))
        self.batches_dir = os.path.join(self.data_dir_full, 'batches')
        self.metadata_path = os.path.join(self.data_dir_full, 'metadata.json')
        self.log_path = os.path.join(os.path.dirname(config_path), self.config.get('log_file', 'hgr.log'))
        
        # Create directories
        os.makedirs(self.data_dir_full, exist_ok=True)
        os.makedirs(self.batches_dir, exist_ok=True)
        
        # Add file handler
        file_handler = logging.FileHandler(self.log_path, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
        self.logger.addHandler(file_handler)
        
        # Add console handler
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        console.setFormatter(formatter)
        self.logger.addHandler(console)
        
        # 初始化处理器和写入器
        self.processor = HGRProcessor(self.config, self.logger)
        self.writer = HGRWriter(
            data_dir=self.data_dir_full,
            summary_filename=self.config.get('summary_filename'),
            batch_template=self.config.get('batch_filename_template'),
            logger=self.logger
        )
        
        self.logger.info("🔧 HGRMain 初始化完成")

    def load_config(self) -> Dict:
        """
        加载配置文件，如果不存在使用默认值
        """
        config = DEFAULT_CONFIG.copy()
        if not os.path.exists(self.config_path):
            self.logger.warning(f"⚠️ 配置文件不存在: {self.config_path}, 使用默认配置")
            return config
        
        cp = configparser.ConfigParser()
        try:
            cp.read(self.config_path, encoding='utf-8')
            for key in config:
                if 'DEFAULT' in cp and key in cp['DEFAULT']:
                    config[key] = cp['DEFAULT'][key]
        except Exception as e:
            self.logger.warning(f"⚠️ 配置文件读取错误: {str(e)}, 使用默认配置")
        
        return config

    def get_full_path(self, path_str: str) -> str:
        """
        处理相对路径（相对于config.ini所在目录）
        """
        if os.path.isabs(path_str):
            return path_str
        config_dir = os.path.dirname(os.path.abspath(self.config_path))
        # config.ini 在 skills/HGRS/，所以 data 就在 skills/HGRS/data
        return os.path.abspath(os.path.join(config_dir, path_str))

    def load_metadata(self) -> Dict:
        """
        加载metadata（自动去重）
        """
        if not os.path.exists(self.metadata_path):
            self.logger.info("ℹ️ metadata.json 不存在，创建新文件")
            return {
                'latest_year': None,
                'latest_batch': None,
                'latest_title': None,
                'processed_batches': [],
                'updated_at': None,
            }
        
        try:
            with open(self.metadata_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 去重处理，防止历史重复数据导致问题
            processed = data.get('processed_batches', [])
            unique_count = len(set(processed))
            if len(processed) != unique_count:
                self.logger.warning(f"⚠️ metadata去重: {len(processed)} → {unique_count}")
                data['processed_batches'] = list(set(processed))
            self.logger.info(f"📂 加载metadata: 已处理 {unique_count} 个批次，最新 {data.get('latest_year')}年第{data.get('latest_batch')}批")
            return data
        except Exception as e:
            self.logger.error(f"❌ metadata加载失败: {str(e)}")
            return {
                'latest_year': None,
                'latest_batch': None,
                'latest_title': None,
                'processed_batches': [],
                'updated_at': None,
            }

    def save_metadata(self, metadata: Dict) -> bool:
        """
        保存metadata
        """
        from datetime import datetime
        metadata['updated_at'] = datetime.now().isoformat()
        try:
            with open(self.metadata_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
            self.logger.info("✅ metadata保存成功")
            return True
        except Exception as e:
            self.logger.error(f"❌ metadata保存失败: {str(e)}")
            return False

    @staticmethod
    def parse_batch_title(title: str) -> Optional[Tuple[int, int]]:
        """
        从标题解析年份和批次号
        返回 (year, batch) 或 None
        """
        # 格式: "中国人类遗传资源行政许可事项2026年第11批审批结果公示"
        import re
        match = re.search(r'(\d{4})年第(\d+)批', title)
        if not match:
            return None
        try:
            year = int(match.group(1))
            batch = int(match.group(2))
            return (year, batch)
        except Exception:
            return None

    def filter_new_batches(self, items: List[Dict], metadata: Dict) -> List[Dict]:
        """
        筛选需要处理的新批次
        按批次号从小到大排序，未处理的就是新批次
        """
        processed = set(metadata.get('processed_batches', []))
        new_items = []
        
        for item in items:
            title = item.get('title', '')
            parsed = self.parse_batch_title(title)
            if not parsed:
                self.logger.debug(f"⚠️ 跳过非公示条目: {title}")
                continue
            
            year, batch = parsed
            key = f"{year}-{batch}"
            if key not in processed:
                item['_year'] = year
                item['_batch'] = batch
                item['_key'] = key
                new_items.append(item)
        
        # 按批次号从小到大排序（先处理旧的，最后处理最新的）
        new_items.sort(key=lambda x: (x['_year'], x['_batch']))
        self.logger.info(f"ℹ️ 筛选结果: 总共 {len(items)} 条，其中 {len(new_items)} 个新批次需要处理")
        return new_items

    def process_new_batch(self, item: Dict, cookies: Dict) -> Optional[Dict]:
        """
        处理单个新批次
        """
        year = item['_year']
        batch = item['_batch']
        title = item['title']
        pdf_url = item['pdf_url']
        
        self.logger.info(f"\n{'='*60}")
        self.logger.info(f"▶️ 开始处理: {title} (PDF: {pdf_url})")
        
        # 下载 → 提取
        result = self.processor.process_pdf(
            pdf_url=pdf_url,
            cookies=cookies,
            save_dir=self.batches_dir,
            year=year,
            batch=batch,
            title=title
        )
        
        if not result:
            self.logger.error(f"❌ 处理失败: {title}")
            return None
        
        # 写入Excel
        write_result = self.writer.process_new_batch(result, self.batches_dir)
        result['write_result'] = write_result
        
        self.logger.info(f"✅ 处理完成: {title}, {result['total_count']} 条记录")
        return result

    def generate_briefing(self, metadata: Dict, processed: List[Dict], errors: List[Dict]) -> str:
        """
        生成运行简报
        """
        from datetime import datetime
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        total_processed = len(metadata.get('processed_batches', []))
        new_processed = len(processed)
        total_errors = len(errors)
        
        total_records = sum(r.get('total_count', 0) for r in processed)
        by_category = {}
        for r in processed:
            for cat, rows in r.get('data', {}).items():
                if rows:
                    by_category[cat] = by_category.get(cat, 0) + len(rows)
        
        lines = []
        lines.append("# 📊 HGR 爬取简报")
        lines.append(f"**运行时间**: {now}")
        lines.append("")
        lines.append(f"本次处理: **{new_processed} 个批次**，新增 **{total_records} 条记录**")
        if total_errors > 0:
            lines.append(f"❌ 失败: **{total_errors} 个批次**（详见日志）")
        lines.append("")
        lines.append("**新增记录按类别统计**:")
        for cat in ['采集审批', '保藏审批', '国际科学研究合作审批', '材料出境证明']:
            cnt = by_category.get(cat, 0)
            if cnt > 0:
                lines.append(f"- {cat}: **{cnt} 条**")
        lines.append("")
        lines.append(f"**累计**: 已处理 **{total_processed} 个批次**")
        if metadata.get('latest_year') and metadata.get('latest_batch'):
            lines.append(f"**最新**: {metadata['latest_year']}年第{metadata['latest_batch']}批 「{metadata.get('latest_title', '')}」")
        lines.append("")
        lines.append("**输出文件**:")
        summary_path = os.path.join(self.data_dir_full, self.config.get('summary_filename'))
        lines.append(f"- 汇总文件: `{summary_path}`")
        if processed:
            lines.append(f"- 独立文件: `{self.batches_dir}/{self.config.get('batch_filename_template').format(year=processed[-1]['year'], batch=processed[-1]['batch'])}`")
        
        return '\n'.join(lines)

    def run_full(self, items: List[Dict], cookies: Dict) -> Dict:
        """
        完整运行流程（全量或增量）
        items: 从页面获取的列表，每个item要有 title 和 pdf_url
        cookies: 浏览器cookie字典
        """
        # 加载metadata
        metadata = self.load_metadata()
        # 筛选新批次
        new_items = self.filter_new_batches(items, metadata)
        
        if not new_items:
            self.logger.info("✅ 没有新批次需要处理，退出")
            return {
                'success': True,
                'has_new': False,
                'briefing': self.generate_briefing(metadata, [], []),
                'metadata': metadata,
            }
        
        # 逐个处理
        processed = []
        errors = []
        
        for item in new_items:
            result = self.process_new_batch(item, cookies)
            if result:
                processed.append(result)
                # 更新metadata
                metadata['processed_batches'].append(item['_key'])
                metadata['latest_year'] = item['_year']
                metadata['latest_batch'] = item['_batch']
                metadata['latest_title'] = item['title']
                # 每成功一个批次立即保存，防止崩溃导致重复处理
                self.save_metadata(metadata)
            else:
                errors.append(item)
        
        # 生成简报
        briefing = self.generate_briefing(metadata, processed, errors)
        
        self.logger.info("\n" + "="*60)
        self.logger.info("🏁 运行完成")
        if processed:
            self.logger.info(f"✅ 成功处理 {len(processed)} 个批次，新增 {sum(r['total_count'] for r in processed)} 条记录")
        if errors:
            self.logger.warning(f"❌ {len(errors)} 个批次处理失败")
        
        return {
            'success': True,
            'has_new': len(processed) > 0,
            'processed': processed,
            'errors': errors,
            'briefing': briefing,
            'metadata': metadata,
        }


def main():
    """
    命令行入口
    用法: python hgr_main.py --items-json items.json --cookies-json cookies.json
    items.json格式: [{"title": "...", "pdf_url": "..."}, ...]
    cookies.json格式: {"name": "value", ...}
    """
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--items-json', required=True, help='批次列表JSON文件')
    parser.add_argument('--cookies-json', required=True, help='Cookie JSON文件')
    parser.add_argument('--config', default='../config.ini', help='配置文件路径')
    args = parser.parse_args()
    
    # 读取items
    with open(args.items_json, 'r', encoding='utf-8') as f:
        items = json.load(f)
    
    # 读取cookies
    with open(args.cookies_json, 'r', encoding='utf-8') as f:
        cookies = json.load(f)
    
    hgr = HGRMain(config_path=args.config)
    result = hgr.run_full(items, cookies)
    
    print("\n" + "="*60)
    print(result['briefing'])
    print("="*60)
    
    return 0 if result['success'] else 1


if __name__ == '__main__':
    sys.exit(main())
