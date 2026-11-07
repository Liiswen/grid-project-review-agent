# -*- coding: utf-8 -*-
"""
配置文件 - 修改此文件中的 API Key 等配置
"""

# 火山引擎 · 豆包配置（同时支持文本和视觉审查）
import os
import sys
try:
    from local_settings import VOLC_API_KEY
except ImportError:
    VOLC_API_KEY = ""
VOLC_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
VOLC_VISION_MODEL = "doubao-seed-2-0-lite-260428"
MAX_UPLOAD_BYTES = 1024 * 1024 * 1024
# ZIP 解压、关键页预览和审查输出仍需占用空间；保留余量以避免收完文件后失败。
MIN_FREE_DISK_BYTES = 3 * 1024 * 1024 * 1024

# 文件路径
RESOURCE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
RUNTIME_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else RESOURCE_DIR
BASE_DIR = RESOURCE_DIR
CHECKPOINTS_FOLDER = os.path.join(BASE_DIR, "检查要点")
EXCEL_TEMPLATE = os.path.join(BASE_DIR, "第二季度规范性检查问题清单模板.xlsx")
UPLOAD_FOLDER = os.path.join(RUNTIME_DIR, "uploads")
UPLOAD_TMP_FOLDER = os.path.join(RUNTIME_DIR, "upload_tmp")
EXTRACT_FOLDER = os.path.join(RUNTIME_DIR, "extracted")
OUTPUT_FOLDER = os.path.join(RUNTIME_DIR, "output")
KNOWLEDGE_FOLDER = EXTRACT_FOLDER

# 支持的文件格式
SUPPORTED_EXTENSIONS = ('.pdf', '.doc', '.docx', '.xlsx', '.wps', '.ofd')
MAX_NESTED_ARCHIVE_BYTES = 1024 * 1024 * 1024
# 招投标资料常把技术规范、报价明细等再次打包。支持有限递归，同时限制层数和数量，
# 防止异常压缩包无限展开。
MAX_NESTED_ARCHIVE_DEPTH = 3
MAX_NESTED_ARCHIVE_COUNT = 100

# 审查性能控制：候选材料先由本地索引和模型匹配收敛，再送入逐条审查。
# 金额、时序、分包等检查天然需要多份交叉证据；候选过少会把已上传的关键文件排除。
MAX_CANDIDATES_PER_CHECKPOINT = 8
REVIEW_CONCURRENCY = 2
API_REQUEST_TIMEOUT_SECONDS = 75
API_REQUEST_RETRIES = 2
MAX_VISION_FILES_PER_CHECKPOINT = 1

# 统一证据管线：扫描 PDF 不再仅抽样传图，而是先进行全页 OCR 并写入页级索引。
# 上限用于防止异常超长附件耗尽任务资源；超过上限会明确记录未 OCR 的页码范围。
ENABLE_FULL_SCAN_OCR = True
OCR_RENDER_DPI = 160
MAX_OCR_PAGES_PER_FILE = 300

