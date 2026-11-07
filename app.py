# -*- coding: utf-8 -*-
"""
Flask Web 应用 - 规范性审查系统
"""
import os
import uuid
import threading
import hashlib
import shutil
import zipfile
import tempfile
from werkzeug.exceptions import HTTPException
from urllib.parse import urlencode
from werkzeug.utils import secure_filename

from flask import (
    Flask, Request as FlaskRequest, render_template, request, jsonify,
    send_file
)

from agent_core import ReviewAgent
from excel_writer import fill_excel_template, generate_simplified_report
import fitz
import config


class UploadRequest(FlaskRequest):
    """仅将 multipart 上传文件的临时流放到 D 盘，不影响 openpyxl 等解析库的临时文件。"""
    def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
        os.makedirs(config.UPLOAD_TMP_FOLDER, exist_ok=True)
        return tempfile.TemporaryFile(mode='w+b', dir=config.UPLOAD_TMP_FOLDER)


app = Flask(
    __name__,
    template_folder=os.path.join(config.RESOURCE_DIR, 'templates'),
)
app.request_class = UploadRequest
app.config['MAX_CONTENT_LENGTH'] = config.MAX_UPLOAD_BYTES  # 1GB 上传限制
# 大 ZIP 的 multipart 请求会先落到 tempfile；上传流由 UploadRequest 改到 D 盘。
os.makedirs(config.UPLOAD_TMP_FOLDER, exist_ok=True)


@app.errorhandler(413)
def request_entity_too_large(error):
    """让超大上传也返回 JSON，避免前端把 HTML 错误页当 JSON 解析。"""
    return jsonify({'error': '上传文件超过 1GB 限制'}), 413


@app.errorhandler(OSError)
def upload_storage_error(error):
    """multipart 临时文件或 uploads 目录写满时，Werkzeug 会在读取 request.files 前抛 OSError。"""
    app.logger.exception('上传文件读写失败')
    return jsonify({'error': f'上传文件读写失败：{error}。请检查系统临时目录和 D 盘剩余空间。'}), 507


@app.errorhandler(Exception)
def unexpected_api_error(error):
    """所有 API 异常均返回 JSON，避免前端把 Flask 的 HTML 错误页当作上传失败。"""
    if isinstance(error, HTTPException):
        return error
    app.logger.exception('未处理的请求异常')
    return jsonify({'error': f'服务器处理上传请求失败：{type(error).__name__}: {error}'}), 500

# 确保目录存在
os.makedirs(config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(config.EXTRACT_FOLDER, exist_ok=True)
os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)

# 存储任务状态
task_status = {}


@app.route('/')
def index():
    """主页"""
    # 本地配置只在服务端作为网页留空时的后备 Key 使用，绝不回填到浏览器页面。
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传 ZIP 文件"""
    if 'file' not in request.files:
        return jsonify({'error': '未选择文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400

    if not file.filename.lower().endswith('.zip'):
        return jsonify({'error': '仅支持 ZIP 格式文件'}), 400

    # 解压、生成关键页图片和输出结果仍需空间。历史重复上传会让保存成功后的审查失败。
    try:
        free_bytes = shutil.disk_usage(config.UPLOAD_FOLDER).free
    except OSError:
        free_bytes = config.MIN_FREE_DISK_BYTES
    if free_bytes < config.MIN_FREE_DISK_BYTES:
        return jsonify({
            'error': '磁盘可用空间不足，无法安全接收并解压项目资料。请先清理历史上传文件或审查输出后重试。'
        }), 507

    # secure_filename 会把中文名清空，项目资料通常使用中文名；仅保留文件名部分并
    # 用任务 ID 隔离，既保留原始显示名，也避免路径穿越。
    original_filename = os.path.basename(file.filename).strip()
    filename = secure_filename(original_filename) or 'project_materials.zip'
    if not filename.lower().endswith('.zip'):
        filename += '.zip'
    task_id = str(uuid.uuid4())[:8]
    save_path = os.path.join(config.UPLOAD_FOLDER, f'{task_id}_{filename}')
    try:
        file.save(save_path)
        if not zipfile.is_zipfile(save_path):
            os.remove(save_path)
            return jsonify({'error': '所选文件不是有效的 ZIP 压缩包'}), 400
    except OSError as exc:
        if os.path.exists(save_path):
            os.remove(save_path)
        return jsonify({'error': f'ZIP 保存失败：{exc}。请检查 uploads 目录的磁盘空间和写入权限。'}), 507

    return jsonify({
        'task_id': task_id,
        'filename': original_filename,
        'filepath': save_path,
        'message': '文件上传成功'
    })


@app.route('/api/review', methods=['POST'])
def start_review():
    """启动审查任务"""
    data = request.get_json()
    task_id = data.get('task_id')
    selected_stages = data.get('stages', [])
    api_key = data.get('api_key', '')

    if not task_id:
        return jsonify({'error': '缺少 task_id'}), 400

    # 找到对应的 ZIP 文件
    zip_path = None
    for f in os.listdir(config.UPLOAD_FOLDER):
        if f.startswith(task_id):
            zip_path = os.path.join(config.UPLOAD_FOLDER, f)
            break

    if not zip_path:
        return jsonify({'error': '文件未找到'}), 404

    # 使用提供的 API Key 或配置中的
    key = api_key or config.VOLC_API_KEY
    if not key or 'your-volc' in key.lower():
        return jsonify({'error': '请先设置火山引擎 API Key'}), 400

    # 初始化任务状态
    task_status[task_id] = {
        'status': 'running',
        'progress': 0,
        'message': '正在初始化...',
        'results': None,
        'excel_path': None,
        'error': None,
        'file_infos': None,
        'manual_review_status': {},  # {seq: 'pending'|'confirmed'} 手动复核状态
        'manual_review_confirmed': False,  # 是否已确认复核完毕
        'api_key': key,
    }

    # 在后台线程中执行审查
    def run_review_task():
        try:
            agent = ReviewAgent(api_key=key)
            available_stages = ReviewAgent.available_stages()
            selected = [s for s in selected_stages if s in available_stages]
            agent.use_stage_rules(selected)

            def update_progress(msg, pct):
                task_status[task_id]['message'] = msg
                task_status[task_id]['progress'] = pct

            results, excel_path, report_path, file_infos = agent.run_review(
                zip_path,
                selected_stages=selected or available_stages,
                progress_callback=update_progress,
                task_id=task_id
            )

            task_status[task_id]['status'] = 'completed'
            task_status[task_id]['progress'] = 100
            task_status[task_id]['message'] = '审查完成'
            task_status[task_id]['results'] = results
            task_status[task_id]['excel_path'] = excel_path
            task_status[task_id]['file_infos'] = file_infos

        except Exception as e:
            task_status[task_id]['status'] = 'error'
            task_status[task_id]['error'] = str(e)
            task_status[task_id]['message'] = f'审查失败: {str(e)}'

    thread = threading.Thread(target=run_review_task)
    thread.daemon = True
    thread.start()

    return jsonify({
        'task_id': task_id,
        'message': '审查任务已启动'
    })


@app.route('/api/status/<task_id>')
def get_status(task_id):
    """获取审查任务状态"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    status = task_status[task_id]
    response = {
        'status': status['status'],
        'progress': status['progress'],
        'message': status['message'],
        'error': status.get('error')
    }

    if status['status'] == 'completed' and status.get('results'):
        # 返回结果摘要
        results = status['results']
        file_infos = status.get('file_infos', [])

        # 修补 empty matched_file：用 _match_files_to_doc 填充
        if file_infos:
            for r in results:
                if not r.get('matched_file'):
                    doc_name = r.get('doc_name', '')
                    matched = _match_files_to_doc(doc_name, '', file_infos)
                    if matched and matched[0].get('filename'):
                        r['matched_file'] = matched[0]['filename']
                        r['matched_file_source'] = 'system_fallback'

        total = len(results)
        qualified = sum(1 for r in results if r['result'] == '合格')
        unqualified = sum(1 for r in results if r['result'] == '不合格')
        unknown = sum(1 for r in results if r['result'] == '无法判断')
        na = sum(1 for r in results if r['result'] == '不适用')

        response['summary'] = {
            'total': total,
            'qualified': qualified,
            'unqualified': unqualified,
            'unknown': unknown,
            'na': na
        }
        response['results'] = results
        response['excel_filename'] = os.path.basename(status.get('excel_path', ''))

    return jsonify(response)


@app.route('/api/download/<task_id>/<file_type>')
def download_result(task_id, file_type):
    """下载审查结果文件"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    status = task_status[task_id]
    if file_type == 'excel':
        # 重新生成 Excel，排除手工复核确认合格的条目
        results = status.get('results', [])
        file_infos = status.get('file_infos', [])
        manual_review_status = status.get('manual_review_status', {})

        # 获取项目名称
        project_name = ''
        for f in os.listdir(config.UPLOAD_FOLDER):
            if f.startswith(task_id):
                project_name = os.path.splitext(os.path.basename(f))[0]
                break

        excel_path = fill_excel_template(
            results, project_name,
            exclude_confirmed_seqs=set(int(k) for k, v in manual_review_status.items() if v == 'confirmed'),
            selected_stages=None
        )
        if not os.path.exists(excel_path):
            return jsonify({'error': '生成 Excel 失败'}), 500
        return send_file(excel_path, as_attachment=True, download_name=os.path.basename(excel_path))

    elif file_type == 'simplified_excel':
        path = status.get('simplified_excel_path')
        if not path or not os.path.exists(path):
            return jsonify({'error': '文件不存在'}), 404
        return send_file(path, as_attachment=True)
    else:
        return jsonify({'error': '无效的文件类型'}), 400


def _normalize_doc_name(doc_name):
    """标准化 doc_name：去除中文/英文括号及内容，提取核心名称"""
    import re
    # 去除中文括号及内容: "（若有）" "（如有）"
    name = re.sub(r'（[^）]*）', '', doc_name)
    # 去除英文括号及内容
    name = re.sub(r'\([^)]*\)', '', name)
    # 去除换行和多余空格
    name = name.replace('\n', '').strip()
    return name


def _similarity(a, b):
    """计算两个字符串的字符级相似度（0~1）"""
    if not a or not b:
        return 0
    a, b = a.lower(), b.lower()
    # 最长公共子序列长度
    m, n = len(a), len(b)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if a[i - 1] == b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])
    lcs = dp[m][n]
    return lcs / max(m, n)


def _extract_core_name(name):
    """提取文件名中的核心中文名称（去除非中文字符后的纯中文部分）"""
    import re
    # 去掉文件扩展名
    name_no_ext = re.sub(r'\.[^.]+$', '', name)
    # 提取所有中文字符
    chinese_chars = re.findall(r'[\u4e00-\u9fff]+', name_no_ext)
    # 取最长的中文片段作为核心名称
    if chinese_chars:
        return max(chinese_chars, key=len)
    return name_no_ext


def _match_files_to_doc(doc_name, matched_file, file_infos):
    """智能匹配文件，按优先级：LLM返回的文件名 > 标准化doc_name匹配 > 模糊匹配 > 最佳猜测"""
    related = []

    # 优先级1: LLM返回的 matched_file
    if matched_file:
        for fi in file_infos:
            fname = fi.get('filename', '')
            # 精确匹配
            if matched_file == fname:
                related.append({
                    'filename': fname,
                    'file_type': fi.get('file_type', ''),
                    'page_count': len(fi.get('page_image_paths', [])),
                    'match_type': 'exact',
                })
            # 子串匹配
            elif matched_file in fname or fname in matched_file:
                related.append({
                    'filename': fname,
                    'file_type': fi.get('file_type', ''),
                    'page_count': len(fi.get('page_image_paths', [])),
                    'match_type': 'substring',
                })
        if related:
            return related

    # 优先级2: 基于 doc_name 的模糊匹配（降低阈值到0.3）
    norm_doc = _normalize_doc_name(doc_name)
    doc_core = _extract_core_name(doc_name)

    scored = []
    for fi in file_infos:
        fname = fi.get('filename', '')
        ftype = fi.get('file_type', '')
        norm_fname = _normalize_doc_name(fname)
        fname_core = _extract_core_name(fname)

        score = 0

        # 子串匹配
        if norm_doc and (norm_doc in norm_fname or norm_doc in fname):
            score = 0.9
        elif doc_name in fname or fname in doc_name:
            score = 0.85
        # 核心名称匹配（提取纯中文部分比较）
        elif doc_core and fname_core:
            if doc_core == fname_core:
                score = 1.0
            elif doc_core in fname_core or fname_core in doc_core:
                score = 0.8
            else:
                # 字符级相似度（降低阈值到0.3，提高容错）
                sim = _similarity(doc_core, fname_core)
                if sim >= 0.3:
                    score = sim

        if score > 0:
            scored.append((score, fi))

    # 按分数降序排列
    scored.sort(key=lambda x: -x[0])
    if scored:
        for score, fi in scored:
            related.append({
                'filename': fi.get('filename', ''),
                'file_type': fi.get('file_type', ''),
                'page_count': len(fi.get('page_image_paths', [])),
                'match_type': 'fuzzy',
            })
        return related

    # 优先级3: 用 doc_name 做关键词子串匹配，找最可能的文件
    # 从 doc_name 中提取关键词片段做子串匹配
    doc_parts = [p for p in doc_name.replace('/', ' ').replace('、', ' ').split() if len(p) >= 2]
    if not doc_parts:
        doc_parts = [doc_name]

    scored = []
    for fi in file_infos:
        fname = fi.get('filename', '')
        match_count = sum(1 for part in doc_parts if part in fname)
        if match_count > 0:
            scored.append((match_count / len(doc_parts), fi))

    scored.sort(key=lambda x: -x[0])
    if scored:
        best_score, best_fi = scored[0]
        related.append({
            'filename': best_fi.get('filename', ''),
            'file_type': best_fi.get('file_type', ''),
            'page_count': len(best_fi.get('page_image_paths', [])),
            'match_type': 'best_guess',
        })
        return related

    # 兜底: 返回第一个文件（而不是全部），避免"无对应文件"因空列表而触发
    if file_infos:
        fi = file_infos[0]
        related.append({
            'filename': fi.get('filename', ''),
            'file_type': fi.get('file_type', ''),
            'page_count': len(fi.get('page_image_paths', [])),
            'match_type': 'fallback',
        })
    return related


def _related_files_for_result(result, file_infos):
    """返回本检查项实际送审的完整候选证据组，而不是只返回主匹配文件。"""
    keys = list(result.get('candidate_files') or [])
    if result.get('matched_file') and result['matched_file'] not in keys:
        keys.insert(0, result['matched_file'])
    related = []
    seen = set()
    for key in keys:
        for info in file_infos:
            file_key = info.get('file_key', info.get('filename', ''))
            if key not in (file_key, info.get('filename', '')) or file_key in seen:
                continue
            seen.add(file_key)
            related.append({
                'filename': info.get('filename', ''),
                'file_key': file_key,
                'file_type': info.get('file_type', ''),
                'page_count': len(info.get('page_image_paths', [])),
                'match_type': 'candidate',
            })
            break
    return related or _match_files_to_doc(result.get('doc_name', ''), result.get('matched_file', ''), file_infos)


@app.route('/api/manual-review/<task_id>', methods=['GET'])
def get_manual_review_items(task_id):
    """获取需要手动复核的条目（无法判断 + 不合格）"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    status = task_status[task_id]
    if status['status'] != 'completed' or not status.get('results'):
        return jsonify({'error': '审查尚未完成'}), 400

    results = status['results']
    file_infos = status.get('file_infos', [])

    # 筛选需要手动复核的条目
    review_items = []
    for r in results:
        if r['result'] in ('无法判断', '不合格'):
            doc_name = r.get('doc_name', '')
            matched_file = r.get('matched_file', '')
            related_files = _related_files_for_result(r, file_infos)

            seq = r['seq']
            review_status = status.get('manual_review_status', {}).get(str(seq), 'pending')

            review_items.append({
                'seq': seq,
                'stage': r.get('stage', ''),
                'link': r.get('link', ''),
                'check_point': r.get('check_point', ''),
                'doc_name': doc_name,
                'matched_file': matched_file,
                'result': r['result'],
                'issue': r.get('issue', ''),
                'suggestion': r.get('suggestion', ''),
                'evidence': r.get('evidence', ''),
                'related_files': related_files,
                'review_status': review_status,
            })

    return jsonify({
        'items': review_items,
        'total': len(review_items),
        'confirmed_count': sum(
            1 for r in results if r['result'] in ('无法判断', '不合格')
            and status.get('manual_review_status', {}).get(str(r['seq']), 'pending') == 'confirmed'
        ),
        'all_confirmed': status.get('manual_review_confirmed', False),
    })


@app.route('/api/manual-review/<task_id>/confirm', methods=['POST'])
def confirm_manual_review(task_id):
    """手动复核确认某一条目合格"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    data = request.get_json()
    seq = data.get('seq')
    if not seq:
        return jsonify({'error': '缺少 seq 参数'}), 400

    status = task_status[task_id]
    if 'manual_review_status' not in status:
        status['manual_review_status'] = {}

    status['manual_review_status'][str(seq)] = 'confirmed'

    # 更新对应结果的 result 为"合格"
    if status.get('results'):
        for r in status['results']:
            if r['seq'] == seq:
                r['original_result'] = r['result']
                r['result'] = '合格'
                r['issue'] = r.get('issue', '') + ' [手动复核已确认合格]'
                break

    return jsonify({'message': '复核确认成功', 'seq': seq})


@app.route('/api/manual-review/<task_id>/finish', methods=['POST'])
def finish_manual_review(task_id):
    """完成手动复核，标记所有待复核项为已确认"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    status = task_status[task_id]
    if 'manual_review_status' not in status:
        status['manual_review_status'] = {}

    # 将所有 pending 的改为 confirmed
    results = status.get('results', [])
    for r in results:
        if r['result'] in ('无法判断', '不合格'):
            seq_str = str(r['seq'])
            if status['manual_review_status'].get(seq_str) != 'confirmed':
                status['manual_review_status'][seq_str] = 'confirmed'
                r['original_result'] = r['result']
                r['result'] = '合格'
                r['issue'] = r.get('issue', '') + ' [手动复核已确认合格]'

    status['manual_review_confirmed'] = True

    return jsonify({'message': '所有项已确认'})


@app.route('/api/manual-review/<task_id>/generate-simplified-report', methods=['POST'])
def generate_simplified_report_endpoint(task_id):
    """复核确认后，调用 LLM 生成简化版 XLSX 报告"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    status = task_status[task_id]
    if status['status'] != 'completed' or not status.get('results'):
        return jsonify({'error': '审查尚未完成'}), 400

    if not status.get('manual_review_confirmed'):
        return jsonify({'error': '请先完成手动复核确认'}), 400

    results = status['results']
    project_name = ''
    for r in results:
        if r.get('project_name'):
            project_name = r['project_name']
            break

    # 从 task_status 中获取项目名称
    # 查找对应的 ZIP 文件名作为项目名
    zip_path = None
    for f in os.listdir(config.UPLOAD_FOLDER):
        if f.startswith(task_id):
            zip_path = os.path.join(config.UPLOAD_FOLDER, f)
            break
    if zip_path:
        project_name = os.path.splitext(os.path.basename(zip_path))[0]

    # 生成简化版报告
    try:
        from reviewer import Reviewer
        reviewer = Reviewer(api_key=status.get('api_key', ''))
        excel_path = generate_simplified_report(results, project_name, reviewer=reviewer)
        status['simplified_excel_path'] = excel_path
        return jsonify({
            'message': '简化报告生成成功',
            'filename': os.path.basename(excel_path)
        })
    except Exception as e:
        return jsonify({'error': f'生成简化报告失败: {str(e)}'}), 500


@app.route('/api/file-image/<task_id>/<path:img_rel_path>')
def serve_file_image(task_id, img_rel_path):
    """提供页面图片供前端查看"""
    images_dir = os.path.join(config.EXTRACT_FOLDER, f'images_{task_id}')
    img_path = os.path.join(images_dir, img_rel_path)
    # 安全检查：确保路径在 images_dir 下
    img_path = os.path.normpath(img_path)
    if not img_path.startswith(os.path.normpath(images_dir)):
        return jsonify({'error': '无效的路径'}), 403
    if not os.path.exists(img_path):
        return jsonify({'error': '图片不存在'}), 404
    return send_file(img_path)


def _find_task_file(task_id, filename):
    """按任务内的原始文件名查找文件信息，避免把文件名当作磁盘路径使用。"""
    if task_id not in task_status:
        return None
    for file_info in task_status[task_id].get('file_infos', []):
        if filename in (file_info.get('file_key'), file_info.get('filename')):
            return file_info
    return None


@app.route('/api/file-page/<task_id>')
def serve_file_page(task_id):
    """按需渲染 PDF 的任意一页，供手动复核查看完整文件。"""
    filename = request.args.get('filename', '')
    try:
        page_number = int(request.args.get('page', '0'))
    except ValueError:
        return jsonify({'error': '页码无效'}), 400

    file_info = _find_task_file(task_id, filename)
    if not file_info:
        return jsonify({'error': '文件未找到'}), 404
    if file_info.get('extension', '').lower() != '.pdf':
        return jsonify({'error': '该文件不是 PDF'}), 400

    source_path = file_info.get('preview_source_path', '') or file_info.get('filepath', '')
    if not source_path or not os.path.exists(source_path):
        return jsonify({'error': '原始文件已不可用'}), 404

    images_dir = os.path.join(config.EXTRACT_FOLDER, f'images_{task_id}', 'full_pages')
    os.makedirs(images_dir, exist_ok=True)
    source_id = hashlib.sha1(filename.encode('utf-8')).hexdigest()[:16]
    image_path = os.path.join(images_dir, f'{source_id}_page_{page_number + 1}.jpg')

    if not os.path.exists(image_path):
        try:
            with fitz.open(source_path) as document:
                if page_number < 0 or page_number >= len(document):
                    return jsonify({'error': '页码超出范围'}), 404
                pix = document[page_number].get_pixmap(dpi=150, alpha=False)
                pix.save(image_path)
        except Exception as exc:
            return jsonify({'error': f'页面转换失败: {exc}'}), 500

    return send_file(image_path)


@app.route('/api/file-images/<task_id>', methods=['POST'])
def get_file_images(task_id):
    """获取指定文件的全部页面图片URL列表；PDF 页面按需渲染，避免只显示关键页。"""
    if task_id not in task_status:
        return jsonify({'error': '任务不存在'}), 404

    data = request.get_json()
    filename = data.get('filename')
    if not filename:
        return jsonify({'error': '缺少 filename'}), 400

    fi = _find_task_file(task_id, filename)
    if fi:
        if fi.get('extension', '').lower() == '.pdf':
            source_path = fi.get('preview_source_path', '') or fi.get('filepath', '')
            if source_path and os.path.exists(source_path):
                try:
                    with fitz.open(source_path) as document:
                        page_count = len(document)
                    query = urlencode({'filename': filename})
                    img_urls = [
                        f'/api/file-page/{task_id}?{query}&page={page_number}'
                        for page_number in range(page_count)
                    ]
                    return jsonify({
                        'filename': filename,
                        'images': img_urls,
                        'page_numbers': list(range(1, page_count + 1)),
                        'total': page_count,
                        'text': fi.get('text', ''),
                    })
                except Exception:
                    pass

        # 兼容旧任务或非 PDF 文件，继续返回已保存的图片。
        img_paths = fi.get('page_image_paths', [])
        img_urls = [f'/api/file-image/{task_id}/{os.path.basename(p)}' for p in img_paths]
        return jsonify({
            'filename': filename,
            'images': img_urls,
            'page_numbers': list(range(1, len(img_urls) + 1)),
            'total': len(img_urls),
            'text': fi.get('text', ''),
        })

    return jsonify({'error': '文件未找到'}), 404


@app.route('/api/stages')
def get_stages():
    return jsonify({'stages': ReviewAgent.available_stages()})


if __name__ == '__main__':
    print("=" * 50)
    print("  规范性审查系统启动")
    print("=" * 50)
    print(f"  火山引擎·豆包: {'已配置' if config.VOLC_API_KEY else '未配置'}")
    print(f"  访问地址: http://127.0.0.1:5000")
    print("=" * 50)
    app.run(debug=False, host='127.0.0.1', port=5000)
