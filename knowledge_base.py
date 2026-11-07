# -*- coding: utf-8 -*-
"""审查任务的临时本地知识库：保存全文，并按检查项检索相关片段。"""

import os
import re
import shutil
import json
import hashlib


def create_knowledge_dir(base_dir, task_id):
    path = os.path.join(base_dir, f"knowledge_{task_id or 'local'}")
    os.makedirs(path, exist_ok=True)
    return path


def save_document_text(knowledge_dir, filename, text):
    safe_name = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._-]+", "_", filename)
    path = os.path.join(knowledge_dir, safe_name + ".txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text or "")
    return path


def build_evidence_units(file_key, text, chunk_size=2600):
    """将全文规范化为可定位、可检索的证据单元。

    PDF 文本及 OCR 文本均使用“--- 第 N 页 ---”边界；其他格式则按标题/段落
    切块。每块保留原始文件相对路径和页码/区块定位，供模型引用而非臆测。
    """
    text = text or ''
    units = []
    page_parts = re.split(r'(?=--- 第 \d+ 页 ---)', text)
    for part in page_parts:
        match = re.match(r'--- 第 (\d+) 页 ---\s*', part)
        locator = f'第 {match.group(1)} 页' if match else ''
        body = part[match.end():] if match else part
        body = body.strip()
        if not body:
            continue
        for index in range(0, len(body), chunk_size):
            chunk = body[index:index + chunk_size].strip()
            if chunk:
                suffix = f'，第 {index // chunk_size + 1} 段' if len(body) > chunk_size else ''
                units.append({
                    'file_key': file_key,
                    'locator': (locator or f'文本区块 {len(units) + 1}') + suffix,
                    'text': chunk,
                })
    if not units and text.strip():
        for index in range(0, len(text), chunk_size):
            units.append({
                'file_key': file_key,
                'locator': f'文本区块 {index // chunk_size + 1}',
                'text': text[index:index + chunk_size].strip(),
            })
    return units


def save_evidence_index(knowledge_dir, file_key, text):
    """保存一个文件的页级/段级证据 JSON，并返回路径。"""
    os.makedirs(knowledge_dir, exist_ok=True)
    safe_id = hashlib.sha1(file_key.encode('utf-8')).hexdigest()
    path = os.path.join(knowledge_dir, f'{safe_id}.evidence.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(build_evidence_units(file_key, text), f, ensure_ascii=False)
    return path


def cleanup_knowledge_dir(path):
    if path and os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
