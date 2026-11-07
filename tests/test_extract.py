# -*- coding: utf-8 -*-
"""解压功能测试：给一个 ZIP，查看系统实际展开了哪些文件。

用法（在项目根目录执行）：
    python tests/test_extract.py 项目资料.zip
    python tests/test_extract.py 项目资料.zip --out extracted/我的测试目录

只调用 archive_extractor（标准库实现），不会加载豆包模型、PDF 解析等依赖。
解压目录默认保留，便于直接打开查看。
"""
import argparse
import os
import sys
import zipfile
from datetime import datetime

# 测试脚本统一放在 tests/ 下，运行时将项目根目录加入模块搜索路径。
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config
from archive_extractor import extract_supported_documents, _io_path

REASON_TEXT = {
    'unsupported': '不在支持的文件类型白名单内',
    'too_large': '内嵌 ZIP 超过单个大小上限',
    'depth_limit': '内嵌 ZIP 超过最大展开层数',
    'count_limit': '内嵌 ZIP 超过最大展开数量',
    'invalid_nested_zip': '扩展名为 .zip 但内容不是有效压缩包（文件本身已保留）',
    'path_traversal': '路径穿越风险，已跳过',
}


def human_size(num):
    for unit in ('B', 'KB', 'MB', 'GB'):
        if num < 1024 or unit == 'GB':
            return f'{num:.1f} {unit}' if unit != 'B' else f'{num} B'
        num /= 1024


def print_tree(directory):
    print('\n解压后的目录结构：')
    # Windows 中文长路径需要扩展长度前缀，否则深层目录可能遍历不到。
    walk_root = _io_path(directory)
    base_depth = walk_root.rstrip(os.sep).count(os.sep)
    for root, dirs, files in os.walk(walk_root):
        dirs.sort()
        depth = root.rstrip(os.sep).count(os.sep) - base_depth
        indent = '  ' * depth
        print(f'{indent}{os.path.basename(root)}/')
        for name in sorted(files):
            path = os.path.join(root, name)
            print(f'{indent}  {name}  ({human_size(os.path.getsize(path))})')


def main():
    parser = argparse.ArgumentParser(description='测试 ZIP 解压结果')
    parser.add_argument('zip_path', help='待测试的 ZIP 文件路径')
    parser.add_argument('--out', help='解压输出目录；不指定时自动建在 extracted/ 下')
    args = parser.parse_args()

    zip_path = os.path.abspath(args.zip_path)
    if not os.path.isfile(zip_path):
        print(f'文件不存在：{zip_path}')
        return 1
    if not zipfile.is_zipfile(zip_path):
        print(f'不是有效的 ZIP 压缩包：{zip_path}')
        return 1

    if args.out:
        out_dir = os.path.abspath(args.out)
    else:
        # 测试目录不拼接项目名，避免目录本身过长；项目内部路径才是长路径的主要来源。
        out_dir = os.path.join(
            config.EXTRACT_FOLDER,
            f'test_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
        )
    os.makedirs(out_dir, exist_ok=True)

    print(f'源文件：{zip_path}  ({human_size(os.path.getsize(zip_path))})')
    print(f'解压到：{out_dir}')

    report = extract_supported_documents(zip_path, out_dir)
    extracted = report['extracted']
    skipped = report['skipped']
    nested = report['nested']

    print(f'\n共展开 {len(extracted)} 个受支持文件，跳过 {len(skipped)} 个条目，'
          f'展开内嵌 ZIP {len(nested)} 个。')

    if extracted:
        by_ext = {}
        for item in extracted:
            ext = os.path.splitext(item['relative_path'])[1].lower() or '(无扩展名)'
            stat = by_ext.setdefault(ext, {'count': 0, 'size': 0})
            stat['count'] += 1
            stat['size'] += item['size']
        print('按类型统计：')
        for ext in sorted(by_ext):
            stat = by_ext[ext]
            print(f'  {ext:<8} {stat["count"]:>3} 个  {human_size(stat["size"])}')

    if nested:
        print('\n展开的内嵌 ZIP：')
        for item in nested:
            print(f'  源：{os.path.basename(item["source"])}')
            print(f'  目录：{os.path.relpath(item["dir"], out_dir)}')

    if skipped:
        print('\n未展开的条目：')
        for item in skipped:
            reason = REASON_TEXT.get(item['reason'], item['reason'])
            print(f'  {item["name"]}  ->  {reason}')

    print_tree(out_dir)
    print(f'\n解压目录已保留：{out_dir}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
