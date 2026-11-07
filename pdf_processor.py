# -*- coding: utf-8 -*-
"""
PDF 解析模块 v2 - 支持扫描件页面图像提取（由豆包视觉模型直接识别）
"""
import os
import io
import re
import subprocess
import sys
import fitz  # pymupdf
import numpy as np
from PIL import Image
import cv2
import config


_ocr_engine = None


def extract_text_from_pdf(pdf_path):
    """从 PDF 中提取文本内容"""
    try:
        doc = fitz.open(pdf_path)
        full_text = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")
            if text.strip():
                full_text.append(f"--- 第 {page_num + 1} 页 ---\n{text}")
        doc.close()
        result = '\n'.join(full_text)
        return result if result.strip() else ''
    except Exception as e:
        return ''


def is_scanned_pdf(pdf_path):
    """判断PDF是否为扫描件（无可提取文本）"""
    text = extract_text_from_pdf(pdf_path)
    return len(text.strip()) < 50


def extract_text_from_scanned_pdf(pdf_path):
    """对扫描 PDF 逐页 OCR，返回带页码边界的完整可检索文本。"""
    global _ocr_engine
    if not config.ENABLE_FULL_SCAN_OCR:
        return '[扫描件：OCR未启用，需通过页面图片人工核验。]'
    try:
        from rapidocr import RapidOCR
        if _ocr_engine is None:
            _ocr_engine = RapidOCR()
        document = fitz.open(pdf_path)
        limit = min(len(document), config.MAX_OCR_PAGES_PER_FILE)
        pages = []
        for page_index in range(limit):
            pix = document[page_index].get_pixmap(dpi=config.OCR_RENDER_DPI, alpha=False)
            output = _ocr_engine(pix.tobytes('png'))
            lines = getattr(output, 'txts', None) or ()
            page_text = '\n'.join(str(line) for line in lines if str(line).strip())
            pages.append(f'--- 第 {page_index + 1} 页 ---\n{page_text}')
        if len(document) > limit:
            pages.append(f'[OCR仅处理前 {limit} 页；第 {limit + 1} 页至第 {len(document)} 页需通过原件页面人工核验。]')
        document.close()
        return '\n'.join(pages)
    except Exception as exc:
        return f'[扫描PDF OCR失败: {exc}；需通过原件页面人工核验。]'


def extract_text_from_docx(docx_path):
    """从 DOCX 中提取文本"""
    try:
        from docx import Document
        doc = Document(docx_path)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        return f"[DOCX文本提取失败: {str(e)}]"


def extract_text_from_doc(doc_path):
    """从 DOC 中提取文本"""
    try:
        from docx import Document
        doc = Document(doc_path)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception:
        pass
    # 很多业务系统将 RTF 表单以 .doc 扩展名导出。它不是 OLE Word 文档，交给
    # Word COM 容易弹出转换窗口；直接按 RTF 读取更稳定，并能保留审批单表格文字。
    try:
        with open(doc_path, 'rb') as source:
            header = source.read(16)
        if header.lstrip().startswith(b'{\\rtf'):
            return extract_text_from_rtf(doc_path)
    except Exception:
        pass
    # 旧版 DOC 常包含复杂表格/修订，python-docx 无法可靠读取。Windows 环境优先
    # 调用已安装的 Word COM 读取正文。COM 打开损坏/受保护文档可能弹窗卡住，故放在
    # 独立进程并限时；不能让一份 DOC 卡死整次 ZIP 审查。
    try:
        creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        if getattr(sys, 'frozen', False):
            command = [sys.executable, '--doc-reader', os.path.abspath(doc_path)]
        else:
            worker = os.path.join(os.path.dirname(__file__), 'doc_reader_worker.py')
            command = [sys.executable, worker, os.path.abspath(doc_path)]
        completed = subprocess.run(
            command,
            capture_output=True, text=True, encoding='utf-8', errors='replace',
            timeout=25, creationflags=creation_flags,
        )
        if completed.returncode == 0 and completed.stdout.strip():
            return completed.stdout.strip()
        detail = (completed.stderr or completed.stdout or 'Word未返回可读取正文').strip()
        return f"[DOC文件已定位，但Word读取未取得正文: {detail[:300]}；需人工核验签字盖章。]"
    except subprocess.TimeoutExpired:
        return "[DOC文件已定位，但Word读取超时；已跳过该文件，需人工核验签字盖章。]"
    except Exception as exc:
        return f"[DOC文件已定位，但Word转换读取失败: {exc}；需人工打开核验签字盖章。]"


def extract_text_from_rtf(rtf_path):
    """提取 RTF 正文，适用于被命名为 .doc 的审批单/流转单。"""
    try:
        raw = open(rtf_path, 'rb').read().decode('latin-1', errors='replace')
        # RTF 的 Unicode 使用 \\uN，紧随其后的一个备用字符不属于正文。
        def unicode_char(match):
            value = int(match.group(1))
            if value < 0:
                value += 65536
            return chr(value)
        text = re.sub(r'\\u(-?\d+)\??', unicode_char, raw)
        text = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes.fromhex(m.group(1)).decode('cp1252', errors='replace'), text)
        text = re.sub(r'\\(par|line|row)\b', '\n', text)
        text = re.sub(r'\\tab\b|\\cell\b', '\t', text)
        text = re.sub(r'\\[a-zA-Z]+-?\d*\s?', '', text)
        text = text.replace('\\{', '{').replace('\\}', '}').replace('\\\\', '\\')
        text = text.replace('{', '').replace('}', '')
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n\s*\n+', '\n', text).strip()
        return text if text else '[RTF文件已定位，但未提取到可见正文；需人工核验。]'
    except Exception as exc:
        return f'[RTF文本提取失败: {exc}]'


def extract_text_from_wps(wps_path):
    """读取 WPS 文字文档中的 Word 二进制正文流。

    常见 .wps 技术规范书是 OLE 容器，正文以 UTF-16 字符串保存在 WordDocument
    流中；不应仅因扩展名为 WPS 就判定“无法解析”。
    """
    try:
        import olefile
        container = olefile.OleFileIO(wps_path)
        if not container.exists('WordDocument'):
            container.close()
            return '[WPS文件已定位，但未找到WordDocument正文流；需人工核验。]'
        raw = container.openstream('WordDocument').read()
        container.close()
        text = raw.decode('utf-16le', errors='ignore')
        # 去除控制字符和二进制残留，保留段落、制表符及可见中英文/标点。
        text = ''.join(char if (char in '\n\t' or ord(char) >= 32) else ' ' for char in text)
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n\s*\n+', '\n', text).strip()
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        if chinese_chars < 20:
            return '[WPS文件已定位，但未提取到足够正文；需人工核验。]'
        return text
    except Exception as exc:
        return f'[WPS文本提取失败: {exc}]'


def extract_text_from_xlsx(xlsx_path):
    """提取Excel各工作表的有效单元格文字，供计划和清单材料检索。"""
    try:
        from openpyxl import load_workbook
        workbook = load_workbook(xlsx_path, read_only=True, data_only=True)
        parts = []
        for sheet in workbook.worksheets:
            parts.append(f"--- 工作表: {sheet.title} ---")
            for row in sheet.iter_rows(values_only=True):
                values = [str(value).strip() for value in row if value not in (None, '')]
                if values:
                    parts.append(' | '.join(values))
        workbook.close()
        return '\n'.join(parts)
    except Exception as exc:
        return f"[XLSX文本提取失败: {exc}]"


def _ofd_localname(tag):
    """OFD XML 元素带固定命名空间，按本地名匹配。"""
    return tag.rsplit('}', 1)[-1]


def _ofd_page_paths(zf):
    """定位 OFD 容器内逐页 Content.xml 的路径列表。"""
    import xml.etree.ElementTree as ET
    names = set(zf.namelist())
    doc_path = None
    if 'OFD.xml' in names:
        try:
            root = ET.fromstring(zf.read('OFD.xml'))
            for el in root.iter():
                if _ofd_localname(el.tag) == 'DocRoot' and el.text and el.text.strip():
                    doc_path = el.text.strip().lstrip('/')
                    break
        except Exception:
            doc_path = None
    if not doc_path or doc_path not in names:
        candidates = [n for n in names if n.endswith('Document.xml') and '/Signs/' not in n]
        doc_path = candidates[0] if candidates else None
    if not doc_path:
        return []
    doc_dir = doc_path.rsplit('/', 1)[0] + '/' if '/' in doc_path else ''
    page_paths = []
    try:
        doc_root = ET.fromstring(zf.read(doc_path))
        for el in doc_root.iter():
            if _ofd_localname(el.tag) == 'Page':
                base = (el.get('BaseLoc') or '').strip().lstrip('/')
                if not base:
                    continue
                full = base if base in names else doc_dir + base
                if full in names:
                    page_paths.append(full)
    except Exception:
        return []
    if not page_paths:
        page_paths = sorted(n for n in names if re.search(r'Pages/Page_\d+/Content\.xml$', n))
    return page_paths


def extract_text_from_ofd(ofd_path):
    """解析 OFD 版式文档（GB/T 33190）正文。

    OFD 是 ZIP 容器：OFD.xml → Document.xml → 每页 Content.xml，
    文字保存在 TextObject/TextCode 元素中；扫描版没有 TextCode，
    正文是整页图片，返回提示并由 extract_ofd_page_images 取图。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(ofd_path) as zf:
            page_paths = _ofd_page_paths(zf)
            if not page_paths:
                return '[OFD文件已定位，但未找到页面内容；需人工核验。]'
            limit = min(len(page_paths), config.MAX_OCR_PAGES_PER_FILE)
            pages = []
            for page_index, page_path in enumerate(page_paths[:limit]):
                try:
                    page_root = ET.fromstring(zf.read(page_path))
                except Exception:
                    continue
                texts = []
                for el in page_root.iter():
                    if _ofd_localname(el.tag) == 'TextCode' and el.text and el.text.strip():
                        texts.append(el.text.strip())
                if texts:
                    pages.append(f'--- 第 {page_index + 1} 页 ---\n' + '\n'.join(texts))
            if len(page_paths) > limit:
                pages.append(f'[仅处理前 {limit} 页；第 {limit + 1} 页至第 {len(page_paths)} 页需通过原件人工核验。]')
            result = '\n'.join(pages)
            if not result.strip():
                return '[OFD无可提取文本，可能为整页扫描图片；将通过页面图片审查。]'
            return result
    except Exception as exc:
        return f'[OFD文本提取失败: {exc}]'


def extract_ofd_page_images(ofd_path):
    """从 OFD 中提取页面整图（ImageObject 引用的 Res 资源）。

    扫描版 OFD 的每一页通常就是一张整页扫描图；文字版由矢量绘制，
    没有整页图片时返回空列表。返回 [(page_index, image_bytes), ...]。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    images = []
    try:
        with zipfile.ZipFile(ofd_path) as zf:
            names = set(zf.namelist())
            page_paths = _ofd_page_paths(zf)
            if not page_paths:
                return images
            # 资源映射：MultiMedia ID -> MediaFile 路径（DocumentRes.xml / PublicRes.xml）
            media = {}
            for res_name in names:
                if not res_name.endswith(('DocumentRes.xml', 'PublicRes.xml')):
                    continue
                res_dir = res_name.rsplit('/', 1)[0] + '/'
                try:
                    res_root = ET.fromstring(zf.read(res_name))
                except Exception:
                    continue
                for el in res_root.iter():
                    if _ofd_localname(el.tag) != 'MultiMedia':
                        continue
                    media_id = el.get('ID')
                    media_file = None
                    for child in el.iter():
                        if _ofd_localname(child.tag) == 'MediaFile' and child.text:
                            media_file = child.text.strip().lstrip('/')
                            break
                    if media_id and media_file:
                        full = media_file if media_file in names else res_dir + media_file
                        if full in names:
                            media[media_id] = full
            for page_index, page_path in enumerate(page_paths):
                if len(images) >= config.MAX_OCR_PAGES_PER_FILE:
                    break
                try:
                    page_root = ET.fromstring(zf.read(page_path))
                except Exception:
                    continue
                for el in page_root.iter():
                    if _ofd_localname(el.tag) != 'ImageObject':
                        continue
                    media_id = el.get('ResourceID')
                    if media_id not in media:
                        continue
                    raw = zf.read(media[media_id])
                    # 统一转 PNG，避免容器内图片格式（如 tiff/jbig2）不被下游识别
                    try:
                        img = Image.open(io.BytesIO(raw))
                        buf = io.BytesIO()
                        img.convert('RGB').save(buf, format='PNG')
                        images.append((page_index, buf.getvalue()))
                    except Exception:
                        images.append((page_index, raw))
                    break  # 每页只取主图
    except Exception:
        pass
    return images


def detect_ofd_signatures(ofd_path):
    """读取 OFD 电子签章清单（Signs/Signatures.xml），供签章类检查参考。"""
    result = {
        'has_signature': False,
        'has_stamp': False,
        'signature_pages': [],
        'stamp_pages': [],
        'details': ''
    }
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(ofd_path) as zf:
            count = 0
            for sig_file in [n for n in zf.namelist() if n.endswith('Signatures.xml')]:
                try:
                    root = ET.fromstring(zf.read(sig_file))
                except Exception:
                    continue
                for el in root.iter():
                    if _ofd_localname(el.tag) == 'Signature':
                        count += 1
            if count:
                result['has_signature'] = True
                result['has_stamp'] = True
                result['details'] = f'OFD包含 {count} 处电子签章记录（Signs/Signatures.xml）'
            else:
                result['details'] = 'OFD未检测到电子签章记录'
    except Exception:
        result['details'] = 'OFD签章信息解析失败'
    return result


def _ofd_page_count(ofd_path):
    import zipfile
    try:
        with zipfile.ZipFile(ofd_path) as zf:
            return len(_ofd_page_paths(zf))
    except Exception:
        return 0


def detect_signatures_and_stamps(pdf_path, key_pages_only=True):
    """
    检测 PDF 中的签名和盖章（基于OpenCV颜色检测）
    """
    result = {
        'has_signature': False,
        'has_stamp': False,
        'signature_pages': [],
        'stamp_pages': [],
        'details': ''
    }

    try:
        doc = fitz.open(pdf_path)
        details = []

        if key_pages_only:
            page_indexes = sorted(set([0, 1, len(doc) - 2, len(doc) - 1]))
        else:
            page_indexes = list(range(len(doc)))
        for page_num in page_indexes:
            if page_num < 0 or page_num >= len(doc):
                continue
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

            stamp_found, stamp_info = _detect_red_stamp(img_cv)
            if stamp_found:
                result['has_stamp'] = True
                result['stamp_pages'].append(page_num + 1)
                details.append(f"第{page_num + 1}页: {stamp_info}")

            sig_found, sig_info = _detect_signature_area(img_cv)
            if sig_found:
                result['has_signature'] = True
                result['signature_pages'].append(page_num + 1)
                details.append(f"第{page_num + 1}页: {sig_info}")

        doc.close()
        result['details'] = '; '.join(details) if details else '未检测到签名或印章'

    except Exception as e:
        result['details'] = f"检测失败: {str(e)}"

    return result


def _detect_red_stamp(img_cv):
    """检测红色印章"""
    hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)
    lower_red1 = np.array([0, 50, 50])
    upper_red1 = np.array([10, 255, 255])
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    lower_red2 = np.array([160, 50, 50])
    upper_red2 = np.array([180, 255, 255])
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = mask1 + mask2

    kernel = np.ones((5, 5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    stamp_regions = []
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > 500:
            perimeter = cv2.arcLength(contour, True)
            if perimeter > 0:
                circularity = 4 * np.pi * area / (perimeter * perimeter)
                if circularity > 0.3:
                    stamp_regions.append(area)

    if stamp_regions:
        return True, f"检测到 {len(stamp_regions)} 处红色印章区域"
    return False, ""


def _detect_signature_area(img_cv):
    """检测签名区域"""
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 100, 255, cv2.THRESH_BINARY_INV)

    hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)
    lower_blue = np.array([100, 50, 50])
    upper_blue = np.array([130, 255, 255])
    blue_mask = cv2.inRange(hsv, lower_blue, upper_blue)

    combined = cv2.bitwise_or(binary, blue_mask)
    kernel = np.ones((3, 3), np.uint8)
    combined = cv2.morphologyEx(combined, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(combined, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    sig_regions = []
    for contour in contours:
        area = cv2.contourArea(contour)
        if 200 < area < 50000:
            x, y, w, h = cv2.boundingRect(contour)
            aspect_ratio = w / h if h > 0 else 0
            if 1.5 < aspect_ratio < 15:
                sig_regions.append(area)

    if sig_regions:
        return True, f"检测到 {len(sig_regions)} 处疑似签名区域"
    return False, ""


def convert_key_pages_to_images(pdf_path, dpi=180):
    """转换PDF首页、第二页、倒数第二页和末页，用于封面及签章核验。"""
    images = []
    try:
        doc = fitz.open(pdf_path)
        indexes = sorted(set([0, 1, len(doc) - 2, len(doc) - 1]))
        for index in indexes:
            if 0 <= index < len(doc):
                images.append((index, doc[index].get_pixmap(dpi=dpi).tobytes('jpg', jpg_quality=85)))
        doc.close()
    except Exception:
        pass
    return images


def process_file(file_path):
    """
    处理单个文件
    扫描件PDF直接转为页面图像，审查时由LLM进行视觉识别
    """
    ext = os.path.splitext(file_path)[1].lower()
    filename = os.path.basename(file_path)

    result = {
        'filename': filename,
        'filepath': file_path,
        'extension': ext,
        'text': '',
        'is_scanned': False,
        'page_count': 0,
        'page_images': [],
        'signature_stamp': {},
        'file_type': _guess_file_type(filename),
    }

    if ext == '.pdf':
        try:
            document = fitz.open(file_path)
            result['page_count'] = len(document)
            document.close()
        except Exception:
            pass
        # 判断是否为扫描件
        text = extract_text_from_pdf(file_path)
        scanned = is_scanned_pdf(file_path)

        result['is_scanned'] = scanned

        if scanned:
            # OCR 全页保留文本证据；视觉模型仍只用于印章、签名、表格版式等 OCR 不可靠信息。
            result['text'] = extract_text_from_scanned_pdf(file_path)
        else:
            result['text'] = text

        # OpenCV检测（作为补充）
        result['signature_stamp'] = detect_signatures_and_stamps(file_path)

    elif ext == '.docx':
        result['text'] = extract_text_from_docx(file_path)
        result['signature_stamp'] = {'details': 'DOCX暂不支持签名印章检测'}
    elif ext == '.doc':
        result['text'] = extract_text_from_doc(file_path)

    elif ext == '.wps':
        result['text'] = extract_text_from_wps(file_path)
        result['signature_stamp'] = {'details': 'WPS暂不支持签名印章检测'}
    elif ext == '.xlsx':
        result['text'] = extract_text_from_xlsx(file_path)
        result['signature_stamp'] = {'details': 'XLSX暂不支持签名印章检测'}
    elif ext == '.ofd':
        result['text'] = extract_text_from_ofd(file_path)
        result['page_count'] = _ofd_page_count(file_path)
        result['signature_stamp'] = detect_ofd_signatures(file_path)
        page_images = extract_ofd_page_images(file_path)
        if page_images:
            result['page_images'] = [img for _, img in page_images]
        # 无文本层但有整页图片时视为扫描件，交给视觉审查
        has_text = bool(result['text']) and not result['text'].startswith('[')
        result['is_scanned'] = bool(page_images) and not has_text
    else:
        result['text'] = f"[不支持的文件格式: {ext}]"

    return result


def _guess_file_type(filename):
    """根据文件名猜测文件类型"""
    keywords = {
        '可研报告': '可研报告', '可行性研究': '可研报告',
        '可研评审': '可研评审意见', '可研批复': '可研批复文件',
        '中标': '中标通知书', '成交': '中标通知书',
        '合同': '合同', '技术规范': '技术规范书',
        '采购': '采购文件', '招标': '招投标文件',
        '需求规格': '需求规格说明书', '业务需求': '业务需求规格说明书',
        '软件需求': '软件需求规格说明书',
        '概要设计': '概要设计说明书', '详细设计': '详细设计说明书',
        '数据字典': '数据字典', '数据库设计': '数据库设计说明书',
        '集成测试': '集成测试报告', '出厂测试': '出厂测试报告',
        '用户确认测试': '用户确认测试报告', '第三方测试': '第三方测试报告',
        '实施方案': '实施方案', '安全防护': '安全防护方案',
        '用户手册': '用户手册', '管理员手册': '系统管理员手册',
        '部署方案': '系统部署方案', '应急预案': '系统应急预案',
        '快速恢复': '快速恢复方案',
        '建设内容完成确认': '项目建设内容完成确认单',
        '上线试运行申请': '上线试运行申请单',
        '上线试运行验收申请': '上线试运行验收申请单',
        '上线试运行验收意见': '上线试运行验收意见',
        '上线运行结果确认': '项目上线运行结果确认单',
        '上线试运行结果确认': '项目上线运行结果确认单',
        '应用成效确认': '应用成效确认单',
        '验收申请': '验收申请单', '验收方案': '验收方案',
        '竣工总结': '竣工总结报告', '工作报告': '工作报告',
        '技术报告': '技术报告', '用户报告': '用户报告',
        '系统测评': '系统测评报告',
        '验收意见': '验收意见', '验收专家签到': '验收专家签到表',
        '结算报告': '项目结算报告', '结算表': '结算表',
        '决算报告': '项目决算报告', '决算审计': '决算审计报告',
        '转资': '转资文件', '记账凭证': '转资文件',
        '启动会': '启动会材料', '建设方案': '建设方案',
        '签到表': '签到表', '运维方案': '运维方案',
        '运维责任备案': '运维责任备案表',
        '红线指标': '红线指标', '蓝线指标': '蓝线指标',
        '红蓝线': '红蓝线指标验证',
    }
    for kw, ftype in keywords.items():
        if kw in filename:
            return ftype
    return '其他文件'
