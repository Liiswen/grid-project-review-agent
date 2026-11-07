# -*- coding: utf-8 -*-
"""ZIP 解压：只展开受支持的文档，并有限递归内嵌 ZIP。仅依赖标准库。"""
import os
import zipfile
import shutil

import config


def decode_archive_name(member):
    """纠正部分国产系统生成的 GBK ZIP 文件名被 zipfile 按 CP437 解码的问题。"""
    name = member.filename.replace('\\', '/')
    if not (member.flag_bits & 0x800):
        try:
            name = name.encode('cp437').decode('gbk')
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    # 统一分隔符并过滤 Windows 不允许的控制字符，后续仍会做路径穿越校验。
    return '/'.join(part.replace('\x00', '') for part in name.split('/') if part not in ('', '.', '..'))


def _io_path(path):
    """Windows 下改用扩展长度路径，避免目录加长中文文件名超过 MAX_PATH 时写入失败。"""
    if os.name != 'nt':
        return path
    absolute = os.path.abspath(path)
    if absolute.startswith('\\\\?\\'):
        return absolute
    if absolute.startswith('\\\\'):
        return '\\\\?\\UNC\\' + absolute[2:]
    return '\\\\?\\' + absolute


def copy_archive_member(archive, member, target):
    """以纠正后的目标路径写入压缩成员，避免 ZipFile.extract 使用原始乱码文件名。"""
    os.makedirs(_io_path(os.path.dirname(target)), exist_ok=True)
    with archive.open(member, 'r') as source, open(_io_path(target), 'wb') as output:
        shutil.copyfileobj(source, output, length=1024 * 1024)


def extract_supported_documents(archive_path, destination, depth=0, nested_state=None):
    """安全展开支持的文档及有限层数的内嵌 ZIP，兼容 GBK 中文文件名。

    返回本次（含内嵌包）的展开报告：
    extracted / skipped / nested 三个列表，供调用方或测试脚本查看实际结果。
    """
    if nested_state is None:
        nested_state = {'count': 0, 'extracted': [], 'skipped': [], 'nested': []}
    root = os.path.abspath(destination)
    try:
        with zipfile.ZipFile(_io_path(archive_path), 'r') as archive:
            for member in archive.infolist():
                if member.is_dir():
                    continue
                member_name = decode_archive_name(member)
                suffix = os.path.splitext(member_name)[1].lower()
                target = os.path.abspath(os.path.join(destination, member_name))
                if not (target == root or target.startswith(root + os.sep)):
                    nested_state['skipped'].append({'name': member_name, 'reason': 'path_traversal'})
                    continue
                if suffix in config.SUPPORTED_EXTENSIONS:
                    copy_archive_member(archive, member, target)
                    nested_state['extracted'].append({
                        'path': target,
                        'relative_path': os.path.relpath(target, root).replace('\\', '/'),
                        'size': os.path.getsize(_io_path(target)),
                    })
                elif (
                    suffix == '.zip'
                    and depth < config.MAX_NESTED_ARCHIVE_DEPTH
                    and nested_state['count'] < config.MAX_NESTED_ARCHIVE_COUNT
                    and member.file_size <= config.MAX_NESTED_ARCHIVE_BYTES
                ):
                    copy_archive_member(archive, member, target)
                    # 扩展名为 zip 不代表内容一定有效；无效内嵌包不能让整个项目审查失败。
                    if not zipfile.is_zipfile(_io_path(target)):
                        nested_state['skipped'].append({
                            'name': member_name,
                            'reason': 'invalid_nested_zip',
                        })
                        continue
                    nested_state['count'] += 1
                    nested_dir = os.path.join(
                        destination, '__nested__',
                        f'd{depth + 1}_{nested_state["count"]}_{os.path.splitext(os.path.basename(member_name))[0]}'
                    )
                    os.makedirs(_io_path(nested_dir), exist_ok=True)
                    nested_state['nested'].append({'source': target, 'dir': nested_dir})
                    extract_supported_documents(target, nested_dir, depth + 1, nested_state)
                else:
                    if suffix == '.zip':
                        if depth >= config.MAX_NESTED_ARCHIVE_DEPTH:
                            reason = 'depth_limit'
                        elif nested_state['count'] >= config.MAX_NESTED_ARCHIVE_COUNT:
                            reason = 'count_limit'
                        elif member.file_size > config.MAX_NESTED_ARCHIVE_BYTES:
                            reason = 'too_large'
                        else:
                            reason = 'unsupported'
                    else:
                        reason = 'unsupported'
                    nested_state['skipped'].append({'name': member_name, 'reason': reason})
    except zipfile.BadZipFile:
        # 顶层包应在上传接口校验；仅容错跳过审查中发现的异常内嵌包。
        if depth == 0:
            raise
    return nested_state
