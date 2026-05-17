#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票/票据智能整理工具

特点：
1. 本地 Gradio Web UI，无需联网、无需部署。
2. PDF 优先读取矢量文本层；只有无法提取或无法解析时才回退 OCR。
3. 图片/扫描件使用 macOS 原生 Vision 框架 OCR，不下载深度学习模型。
4. 自动重命名归档，并生成「发票报销统计表.xlsx」。

运行：
    python3 invoice_sorter_webui.py
"""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import traceback
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional

import pandas as pd
import pdfplumber
from PIL import Image, ImageDraw, ImageFont


APP_SUPPORT_DIR = Path.home() / "Library" / "Application Support" / "Invoicely"
RATES_CACHE_FILE = APP_SUPPORT_DIR / "rates.json"
SETTINGS_FILE = APP_SUPPORT_DIR / "settings.json"

SUPPORTED_CURRENCIES = ["CNY", "HKD", "USD", "GBP", "EUR", "JPY", "THB", "AED"]

# 离线兜底汇率：以 1 USD 可兑换多少目标货币为单位。
# 网络不可用时仍可继续整理，之后联网打开 App 会自动更新。
FALLBACK_USD_RATES = {
    "USD": 1.0,
    "CNY": 7.20,
    "HKD": 7.80,
    "GBP": 0.79,
    "EUR": 0.92,
    "JPY": 155.0,
    "THB": 36.0,
    "AED": 3.6725,
}

SUPPORTED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}

TOTAL_KEYWORDS = [
    "Total com IVA",
    "Total with VAT",
    "TOTAL DUE",
    "Amount Paid",
    "Grand Total",
    "Balance Due",
    "Total",
]


@dataclass
class OCRLine:
    """Vision OCR 单行结果，bbox 为 Vision 归一化坐标，原点在左下角。"""

    text: str
    x: float
    y: float
    width: float
    height: float


@dataclass
class InvoiceResult:
    date: str
    merchant: str
    amount_original: float
    original_currency: str
    invoice_type: str
    original_name: str
    new_name: str = ""
    category: str = "Other"
    report_amount: int = 0
    report_currency: str = "CNY"

    @property
    def amount_cny(self) -> int:
        """兼容旧代码读取；新逻辑使用 report_amount/report_currency。"""

        return self.report_amount or round(self.amount_original)


@dataclass
class AppSettings:
    filename_format: str = "{yymm}_{seq}_{category}_{currency}{amount}"
    auto_update_rates: bool = True
    annotate_converted: bool = True
    sequence_start: int = 1

    @classmethod
    def load(cls) -> "AppSettings":
        try:
            if SETTINGS_FILE.exists():
                data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
                return cls(
                    filename_format=str(data.get("filename_format") or cls.filename_format),
                    auto_update_rates=bool(data.get("auto_update_rates", True)),
                    annotate_converted=bool(data.get("annotate_converted", True)),
                    sequence_start=int(data.get("sequence_start", 1)),
                )
        except Exception:
            pass
        return cls()

    def save(self) -> None:
        APP_SUPPORT_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(
                {
                    "filename_format": self.filename_format,
                    "auto_update_rates": self.auto_update_rates,
                    "annotate_converted": self.annotate_converted,
                    "sequence_start": self.sequence_start,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )


class RateManager:
    """管理在线汇率缓存。缓存很小，仅保存当天汇率，避免每次处理都联网。"""

    def __init__(self) -> None:
        self.updated_date = ""
        self.rates = dict(FALLBACK_USD_RATES)
        self.source = "离线兜底汇率"
        self.load_cache()

    def load_cache(self) -> None:
        try:
            if not RATES_CACHE_FILE.exists():
                return
            data = json.loads(RATES_CACHE_FILE.read_text(encoding="utf-8"))
            cached = data.get("rates") or {}
            for code in SUPPORTED_CURRENCIES:
                if code in cached:
                    self.rates[code] = float(cached[code])
            self.updated_date = str(data.get("date") or "")
            self.source = str(data.get("source") or "本地缓存")
        except Exception:
            self.rates = dict(FALLBACK_USD_RATES)
            self.source = "离线兜底汇率"

    def save_cache(self) -> None:
        APP_SUPPORT_DIR.mkdir(parents=True, exist_ok=True)
        RATES_CACHE_FILE.write_text(
            json.dumps(
                {
                    "date": self.updated_date,
                    "source": self.source,
                    "rates": {code: self.rates[code] for code in SUPPORTED_CURRENCIES},
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    def update_if_needed(self, force: bool = False) -> tuple[bool, str]:
        today = datetime.now().strftime("%Y-%m-%d")
        if not force and self.updated_date == today:
            return False, f"今日汇率已更新：{self.short_status()}"

        try:
            request = urllib.request.Request(
                "https://open.er-api.com/v6/latest/USD",
                headers={"User-Agent": "Invoicely/1.0"},
            )
            with urllib.request.urlopen(request, timeout=8) as response:
                data = json.loads(response.read().decode("utf-8"))
            rates = data.get("rates") or {}
            for code in SUPPORTED_CURRENCIES:
                if code not in rates:
                    raise RuntimeError(f"在线汇率缺少 {code}")
                self.rates[code] = float(rates[code])
            self.updated_date = today
            self.source = "open.er-api.com"
            self.save_cache()
            return True, f"汇率已更新：{self.short_status()}"
        except Exception as exc:
            return False, f"汇率更新失败，继续使用{self.source}：{exc}"

    def convert(self, amount: float, from_currency: str, to_currency: str) -> float:
        source = normalize_currency_code(from_currency)
        target = normalize_currency_code(to_currency)
        if source == target:
            return float(amount)
        source_rate = self.rates.get(source) or FALLBACK_USD_RATES.get(source)
        target_rate = self.rates.get(target) or FALLBACK_USD_RATES.get(target)
        if not source_rate or not target_rate:
            raise RuntimeError(f"缺少汇率：{source} -> {target}")
        usd_amount = float(amount) / source_rate
        return usd_amount * target_rate

    def short_status(self) -> str:
        date = self.updated_date or "未更新"
        return f"{date} · USD/CNY {self.rates['CNY']:.4g} · USD/HKD {self.rates['HKD']:.4g}"


def normalize_date(raw: str | None) -> str:
    """把 2026年05月09日、2026/05/09、2026-05-09 等格式统一为 YYYY-MM-DD。"""

    if not raw:
        return datetime.today().strftime("%Y-%m-%d")

    raw = raw.strip()
    patterns = [
        r"(?P<y>20\d{2})\s*年\s*(?P<m>\d{1,2})\s*月\s*(?P<d>\d{1,2})\s*日",
        r"(?P<y>20\d{2})[./-](?P<m>\d{1,2})[./-](?P<d>\d{1,2})",
        r"(?P<d>\d{1,2})[./-](?P<m>\d{1,2})[./-](?P<y>20\d{2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, raw)
        if match:
            y = int(match.group("y"))
            m = int(match.group("m"))
            d = int(match.group("d"))
            return f"{y:04d}-{m:02d}-{d:02d}"
    return datetime.today().strftime("%Y-%m-%d")


def sanitize_filename_part(text: str, fallback: str = "未知商家") -> str:
    """清理文件名中的非法字符，并控制长度。"""

    cleaned = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "", text or "").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return (cleaned[:40] or fallback).strip()


def unique_path(path: Path) -> Path:
    """如果目标文件已存在，自动追加序号，避免覆盖已有文件。"""

    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for idx in range(1, 1000):
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法生成不重名文件名：{path.name}")


def extract_pdf_text(path: Path) -> str:
    """读取 PDF 矢量文本层。适用于国内电子发票、火车票等标准 PDF。"""

    chunks: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def parse_china_standard_invoice(text: str, original_name: str) -> Optional[InvoiceResult]:
    """解析中国数电发票/电子普通发票。"""

    if not any(k in text for k in ["电子发票", "发票号码", "价税合计"]):
        return None
    if "电子客票" in text or "12306" in text:
        return None

    date_match = re.search(r"开票日期[:：]?\s*(20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|20\d{2}[./-]\d{1,2}[./-]\d{1,2})", text)
    date = normalize_date(date_match.group(1) if date_match else text)

    amount_patterns = [
        r"[（(]\s*小写\s*[）)]\s*[¥￥]?\s*([0-9,]+(?:\.\d{1,2})?)",
        r"价税合计.*?[¥￥]\s*([0-9,]+(?:\.\d{1,2})?)",
        r"合\s*计\s*[¥￥]\s*[0-9,]+(?:\.\d{1,2})?\s*[¥￥]\s*([0-9,]+(?:\.\d{1,2})?)",
    ]
    amount = first_amount_by_patterns(text, amount_patterns)
    if amount is None:
        return None

    merchant = extract_china_seller_name(text)
    return InvoiceResult(
        date=date,
        merchant=merchant,
        amount_original=amount,
        original_currency="CNY",
        invoice_type="中国标准发票",
        original_name=original_name,
        category=infer_category(merchant + "\n" + text),
    )


def extract_china_seller_name(text: str) -> str:
    """尽量提取“销售方信息”下方的名称。PDF 文本层常会被表格拆散，所以准备多种兜底规则。"""

    match = re.search(r"销售方信息.*?名称[:：]\s*([^\n]+)", text, flags=re.S)
    if match:
        return sanitize_filename_part(match.group(1))

    # 样例中购买方和销售方名称常在同一行：买方公司 销方公司。
    company_names = re.findall(r"[\u4e00-\u9fa5（）()A-Za-z0-9]+(?:有限公司|公司|酒店|中心|商行|店)", text)
    if len(company_names) >= 2:
        return sanitize_filename_part(company_names[1])
    if company_names:
        return sanitize_filename_part(company_names[-1])
    return "未知商家"


def parse_china_train_ticket(text: str, original_name: str) -> Optional[InvoiceResult]:
    """解析中国铁路电子客票。"""

    if not any(k in text for k in ["电子客票", "12306", "中国铁路", "票价"]):
        return None

    date_match = re.search(r"(20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|20\d{2}[./-]\d{1,2}[./-]\d{1,2})", text)
    amount = extract_train_amount(text)
    if amount is None:
        return None

    return InvoiceResult(
        date=normalize_date(date_match.group(1) if date_match else text),
        merchant="中国铁路12306",
        amount_original=amount,
        original_currency="CNY",
        invoice_type="中国标准发票",
        original_name=original_name,
        category="Transport",
    )


def extract_train_amount(text: str) -> Optional[float]:
    """火车票金额常在“￥167.00”和“票价:”附近，避免误把身份证号/票号当金额。"""

    candidates: list[float] = []
    for match in re.finditer(r"[¥￥]\s*([0-9,]+(?:\.\d{1,2})?)", text):
        amount = parse_number(match.group(1))
        if amount is not None and 0 < amount < 10000:
            candidates.append(amount)
    if candidates:
        return candidates[0]

    ticket_window = ""
    price_pos = text.find("票价")
    if price_pos >= 0:
        ticket_window = text[max(0, price_pos - 80) : price_pos + 80]
    match = re.search(r"([0-9,]+(?:\.\d{1,2}))", ticket_window)
    if match:
        return parse_number(match.group(1))
    return None


def first_amount_by_patterns(text: str, patterns: Iterable[str]) -> Optional[float]:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I | re.S)
        if match:
            return parse_number(match.group(1))
    return None


def parse_number(raw: str | None) -> Optional[float]:
    """兼容 1,234.56 和葡语/欧陆格式 1.234,56。"""

    if not raw:
        return None
    value = re.sub(r"[^\d,.\-]", "", raw)
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif "," in value and "." not in value:
        value = value.replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return None


def import_vision_stack():
    """延迟导入 pyobjc 相关模块，方便 Web UI 启动时给出友好错误。"""

    try:
        import Vision
        import Quartz
        from Foundation import NSURL
    except Exception as exc:
        raise RuntimeError(
            "无法导入 macOS Vision/Quartz。请确认已安装 "
            "pyobjc-framework-Vision 和 pyobjc-framework-Quartz。"
        ) from exc
    return Vision, Quartz, NSURL


def vision_ocr_image(image_path: Path) -> list[OCRLine]:
    """使用 macOS Vision 对图片做 OCR。不会下载任何模型权重。"""

    Vision, Quartz, NSURL = import_vision_stack()
    url = NSURL.fileURLWithPath_(str(image_path))
    handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(url, None)
    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    request.setUsesLanguageCorrection_(True)
    try:
        request.setRecognitionLanguages_(["zh-Hans", "en-US", "pt-PT", "ja-JP"])
    except Exception:
        # 较旧系统可能不接受全部语言，保持默认识别也可继续。
        pass

    outcome = handler.performRequests_error_([request], None)
    if isinstance(outcome, tuple):
        ok, err = outcome
        if ok is False and err is not None:
            raise RuntimeError(f"Vision OCR 执行失败：{err}")

    lines: list[OCRLine] = []
    for observation in request.results() or []:
        candidates = observation.topCandidates_(1)
        if not candidates:
            continue
        text = str(candidates[0].string()).strip()
        if not text:
            continue
        box = observation.boundingBox()
        lines.append(OCRLine(text=text, x=box.origin.x, y=box.origin.y, width=box.size.width, height=box.size.height))

    # Vision 原点在左下角。这里按视觉阅读顺序排序：上到下、左到右。
    lines.sort(key=lambda item: (-item.y, item.x))
    return lines


def render_pdf_first_page_to_image(pdf_path: Path, output_path: Path, scale: float = 2.0) -> Path:
    """用 macOS Quartz 把 PDF 首页渲染成 PNG，供 Vision OCR 或 Pillow 标注使用。"""

    _, Quartz, NSURL = import_vision_stack()
    url = NSURL.fileURLWithPath_(str(pdf_path))
    document = Quartz.CGPDFDocumentCreateWithURL(url)
    if not document:
        raise RuntimeError(f"无法读取 PDF：{pdf_path.name}")

    page = Quartz.CGPDFDocumentGetPage(document, 1)
    media_box = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
    width = int(media_box.size.width * scale)
    height = int(media_box.size.height * scale)

    color_space = Quartz.CGColorSpaceCreateDeviceRGB()
    raw = bytearray(width * height * 4)
    context = Quartz.CGBitmapContextCreate(
        raw,
        width,
        height,
        8,
        width * 4,
        color_space,
        Quartz.kCGImageAlphaPremultipliedLast,
    )
    Quartz.CGContextSetRGBFillColor(context, 1, 1, 1, 1)
    Quartz.CGContextFillRect(context, Quartz.CGRectMake(0, 0, width, height))
    Quartz.CGContextScaleCTM(context, scale, scale)
    Quartz.CGContextDrawPDFPage(context, page)

    image = Image.frombytes("RGBA", (width, height), bytes(raw))
    image = image.transpose(Image.Transpose.FLIP_TOP_BOTTOM).convert("RGB")
    image.save(output_path)
    return output_path


def ocr_document(path: Path) -> list[OCRLine]:
    """图片直接 OCR；PDF 先渲染首页再 OCR。"""

    if path.suffix.lower() == ".pdf":
        with tempfile.TemporaryDirectory() as td:
            image_path = Path(td) / "page1.png"
            render_pdf_first_page_to_image(path, image_path)
            return vision_ocr_image(image_path)
    return vision_ocr_image(path)


def parse_foreign_manual(lines: list[OCRLine], original_name: str) -> Optional[InvoiceResult]:
    """解析已在右上角手动写 CNY/HKD/#金额 的国外票据。# 默认视为人民币旧习惯标注。"""

    upper_right = [line for line in lines if line.x > 0.60 and line.y > 0.78]
    text = "\n".join(line.text for line in upper_right)
    match = re.search(r"(?P<currency>HKD|HK\$|CNY|RMB|人民币|#)?\s*(?P<amount>[1-9]\d{1,6})(?:\.00)?\b", text, flags=re.I)
    if not match:
        return None

    full_text = "\n".join(line.text for line in lines)
    date = extract_any_date(full_text)
    merchant = extract_foreign_merchant(lines)
    amount = int(match.group("amount"))
    currency = normalize_currency_code(match.group("currency") or "CNY")
    return InvoiceResult(
        date=date,
        merchant=merchant,
        amount_original=amount,
        original_currency=currency,
        invoice_type="国外手动",
        original_name=original_name,
        category=infer_category(merchant + "\n" + full_text),
    )


def parse_foreign_program(lines: list[OCRLine], original_name: str) -> Optional[InvoiceResult]:
    """解析未手动标注的国外票据，稍后按当前模式统一换算。"""

    full_text = "\n".join(line.text for line in lines)
    amount_foreign, currency = find_foreign_total(full_text)
    if amount_foreign is None:
        return None

    merchant = extract_foreign_merchant(lines)
    return InvoiceResult(
        date=extract_any_date(full_text),
        merchant=merchant,
        amount_original=amount_foreign,
        original_currency=currency,
        invoice_type="国外程序标注",
        original_name=original_name,
        category=infer_category(merchant + "\n" + full_text),
    )


def parse_foreign_program_text(text: str, original_name: str) -> Optional[InvoiceResult]:
    """解析带文本层的国外 PDF，避免对清晰电子 PDF 做不必要 OCR。"""

    amount_foreign, currency = find_foreign_total(text)
    if amount_foreign is None:
        return None

    merchant = extract_foreign_merchant_from_text(text)
    return InvoiceResult(
        date=extract_any_date(text),
        merchant=merchant,
        amount_original=amount_foreign,
        original_currency=currency,
        invoice_type="国外程序标注",
        original_name=original_name,
        category=infer_category(merchant + "\n" + text),
    )


def find_foreign_total(text: str) -> tuple[Optional[float], str]:
    """围绕 Total 类关键词寻找金额和币种。"""

    compact = re.sub(r"[ \t]+", " ", text)
    for keyword in TOTAL_KEYWORDS:
        for match in re.finditer(re.escape(keyword), compact, flags=re.I):
            window = compact[match.start() : match.end() + 180]
            amount_match = re.search(
                r"(?P<symbol>€|EUR|HK\$|HKD|US\$|\$|USD|£|GBP|JPY|¥|CNY|RMB|THB|฿|AED|د\.إ)?\s*"
                r"(?P<amount>\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})|\d+(?:[.,]\d{2})?)",
                window,
                flags=re.I,
            )
            if amount_match:
                amount = parse_number(amount_match.group("amount"))
                currency = infer_currency(amount_match.group("symbol") or window or compact)
                if amount is not None:
                    return amount, currency
    return None, "USD"


def infer_currency(context: str) -> str:
    context_upper = (context or "").upper()
    if "HK$" in context_upper or "HKD" in context_upper:
        return "HKD"
    if "RMB" in context_upper or "CNY" in context_upper or "人民币" in context or "￥" in context:
        return "CNY"
    if "AED" in context_upper or "د.إ" in context or "DIRHAM" in context_upper:
        return "AED"
    if "THB" in context_upper or "฿" in context or "BAHT" in context_upper:
        return "THB"
    if "€" in context or "EUR" in context_upper or "IVA" in context_upper or "PORTUGAL" in context_upper:
        return "EUR"
    if "£" in context or "GBP" in context_upper:
        return "GBP"
    if "JPY" in context_upper or "円" in context or "¥" in context:
        return "JPY"
    if "$" in context or "USD" in context_upper or "US$" in context_upper:
        return "USD"
    return "USD"


def normalize_currency_code(raw: str | None) -> str:
    text = (raw or "").upper().strip()
    if text in {"RMB", "CNY", "人民币", "￥", "#"}:
        return "CNY"
    if text in {"HK$", "HKD"}:
        return "HKD"
    if text in {"US$", "$", "USD"}:
        return "USD"
    if text in {"€", "EUR"}:
        return "EUR"
    if text in {"£", "GBP"}:
        return "GBP"
    if text in {"¥", "JPY"}:
        return "JPY"
    if text in {"฿", "THB"}:
        return "THB"
    if text in {"AED", "د.إ"}:
        return "AED"
    return text if text in SUPPORTED_CURRENCIES else "USD"


def extract_any_date(text: str) -> str:
    match = re.search(
        r"(20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|20\d{2}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2}[./-]20\d{2})",
        text,
    )
    return normalize_date(match.group(1) if match else None)


def extract_foreign_merchant(lines: list[OCRLine]) -> str:
    """国外小票通常头部 1-2 行就是商家名称。过滤掉明显的发票标题和日期行。"""

    skip_words = {"INVOICE", "RECEIPT", "FATURA", "ORIGINAL", "TAX INVOICE"}
    candidates: list[str] = []
    for line in lines[:8]:
        text = line.text.strip()
        if len(text) < 2:
            continue
        upper = text.upper()
        if any(word in upper for word in skip_words):
            continue
        if re.search(r"20\d{2}[./-]\d{1,2}[./-]\d{1,2}", text):
            continue
        candidates.append(text)
        if len(candidates) >= 2:
            break
    return sanitize_filename_part(" ".join(candidates) if candidates else "ForeignReceipt", "ForeignReceipt")


def extract_foreign_merchant_from_text(text: str) -> str:
    """从 PDF 文本层前几行提取国外商家名称。"""

    skip_words = {"INVOICE", "RECEIPT", "FATURA", "ORIGINAL", "TAX INVOICE"}
    candidates: list[str] = []
    for raw_line in text.splitlines()[:12]:
        line = raw_line.strip()
        if len(line) < 2:
            continue
        upper = line.upper()
        if any(word in upper for word in skip_words):
            continue
        if re.search(r"20\d{2}[./-]\d{1,2}[./-]\d{1,2}", line):
            continue
        candidates.append(line)
        if len(candidates) >= 2:
            break
    return sanitize_filename_part(" ".join(candidates) if candidates else "ForeignReceipt", "ForeignReceipt")


def annotate_document_with_label(source: Path, label_text: str, target: Path) -> None:
    """在右上角写入醒目的换算金额。PDF 保持原始页面方向，只叠加标注图。"""

    if source.suffix.lower() == ".pdf":
        _, Quartz, NSURL = import_vision_stack()

        # 读取原始 PDF 首页尺寸
        src_url = NSURL.fileURLWithPath_(str(source))
        document = Quartz.CGPDFDocumentCreateWithURL(src_url)
        if not document:
            raise RuntimeError(f"无法读取 PDF：{source.name}")
        page = Quartz.CGPDFDocumentGetPage(document, 1)
        media_box = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
        page_w = media_box.size.width
        page_h = media_box.size.height

        # 创建输出 PDF（保持原始页面尺寸）
        output_url = NSURL.fileURLWithPath_(str(target))
        data_consumer = Quartz.CGDataConsumerCreateWithURL(output_url)
        out_rect = Quartz.CGRectMake(0, 0, page_w, page_h)
        pdf_context = Quartz.CGPDFContextCreate(data_consumer, out_rect, None)
        Quartz.CGPDFContextBeginPage(pdf_context, None)

        # 1) 直接绘制原始 PDF 页面（Quartz 原生坐标系，方向完全正确）
        Quartz.CGContextDrawPDFPage(pdf_context, page)

        # 2) 用 PIL 生成标注小图，再叠加到 PDF 右上角
        label_image = _make_label_image(label_text, page_w)
        label_w, label_h = label_image.size

        with tempfile.TemporaryDirectory() as td:
            label_png = Path(td) / "label.png"
            label_image.save(label_png)

            label_url = NSURL.fileURLWithPath_(str(label_png))
            img_src = Quartz.CGImageSourceCreateWithURL(label_url, None)
            cg_label = Quartz.CGImageSourceCreateImageAtIndex(img_src, 0, None)

        # PDF 坐标系原点在左下角、Y 向上，"右上角"= 大 x、大 y
        margin = page_w * 0.018
        dest_x = page_w - margin - label_w
        dest_y = page_h - margin - label_h
        dest_rect = Quartz.CGRectMake(dest_x, dest_y, label_w, label_h)

        # 直接绘制标注图到目标位置
        Quartz.CGContextDrawImage(pdf_context, dest_rect, cg_label)

        Quartz.CGPDFContextEndPage(pdf_context)
        Quartz.CGPDFContextClose(pdf_context)
    else:
        image = Image.open(source).convert("RGB")
        annotated = draw_amount_label(image, label_text)
        save_kwargs = {"quality": 95} if target.suffix.lower() in {".jpg", ".jpeg"} else {}
        annotated.save(target, **save_kwargs)


def annotate_document_with_cny(source: Path, amount_cny: int, target: Path) -> None:
    """兼容旧调用。"""

    annotate_document_with_label(source, f"CNY {amount_cny}", target)


def _make_label_image(text: str, page_width: float) -> Image.Image:
    """生成金额标注小 RGBA 图片，用于 PDF 叠加。page_width 为 PDF 点数。"""

    # PDF 页面通常 ~595pt 宽（A4），字号按页面宽度缩放
    font_size = max(16, int(page_width * 0.028))
    try:
        font = ImageFont.truetype("/System/Library/Fonts/PingFang.ttc", font_size)
    except Exception:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)

    # 先测量文字尺寸
    tmp = Image.new("RGBA", (1, 1))
    draw = ImageDraw.Draw(tmp)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    pad = max(4, int(page_width * 0.008))
    line_width = max(2, int(page_width * 0.003))
    img_w = text_w + pad * 2 + line_width * 2
    img_h = text_h + pad * 2 + line_width * 2

    img = Image.new("RGBA", (img_w, img_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, img_w - 1, img_h - 1], fill=(255, 235, 59, 255),
                    outline=(220, 0, 0, 255), width=line_width)
    draw.text((line_width + pad - bbox[0], line_width + pad - bbox[1]),
              text, fill=(0, 0, 0, 255), font=font)
    return img


def _make_cny_label_image(amount_cny: int, page_width: float) -> Image.Image:
    """兼容旧调用。"""

    return _make_label_image(f"CNY {amount_cny}", page_width)


def draw_amount_label(image: Image.Image, text: str) -> Image.Image:
    """用 Pillow 在右上角绘制红框和黄底黑字。"""

    image = image.copy()
    draw = ImageDraw.Draw(image)
    width, height = image.size
    # 字体大小基于图像宽度计算
    font_size = max(80, int(width * 0.04))
    try:
        font = ImageFont.truetype("/System/Library/Fonts/PingFang.ttc", font_size)
    except Exception:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)

    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    pad = int(width * 0.015)
    margin = int(width * 0.025)
    box_w = text_w + pad * 2
    box_h = text_h + pad * 2

    x2 = width - margin
    y1 = margin
    x1 = x2 - box_w
    y2 = y1 + box_h

    line_width = max(6, int(width * 0.004))
    draw.rectangle([x1, y1, x2, y2], fill=(255, 235, 59), outline=(220, 0, 0), width=line_width)
    draw.text((x1 + pad, y1 + pad), text, fill=(0, 0, 0), font=font)
    return image


def draw_cny_label(image: Image.Image, amount_cny: int) -> Image.Image:
    """兼容旧调用。"""

    return draw_amount_label(image, f"CNY {amount_cny}")


def infer_category(text: str) -> str:
    upper = text.upper()
    if any(k in upper for k in ["MEAL", "RESTAURANT", "FOOD", "CAFE", "COFFEE", "DINING", "BAR", "餐", "咖啡", "饭店", "餐厅"]):
        return "F&B"
    if any(k in upper for k in ["HOTEL", "HILTON", "MARRIOTT", "HYATT", "住宿", "KYOTO", "旅馆", "酒店", "INN"]):
        return "Hotel"
    if any(k in upper for k in ["RENT-A-CAR", "CAR RENTAL", "PORTAGENS", "TOLL", "TRAIN", "RAIL", "TAXI", "UBER", "FLIGHT", "AIRLINE", "PARKING", "火车", "铁路", "12306", "租车", "交通"]):
        return "Transport"
    if any(k in upper for k in ["OFFICE", "STATIONERY", "SUPPLIES", "PRINTER", "文具", "办公"]):
        return "Office"
    if any(k in upper for k in ["TELECOM", "MOBILE", "PHONE", "DATA", "INTERNET", "BROADBAND", "电话", "宽带", "流量"]):
        return "Telecom"
    if any(k in upper for k in ["SOFTWARE", "SAAS", "CLOUD", "SUBSCRIPTION", "OPENAI", "GITHUB", "APPLE", "GOOGLE", "MICROSOFT"]):
        return "Software"
    if any(k in upper for k in ["VISA", "AIRPORT", "TRAVEL", "TRIP", "签证", "机场"]):
        return "Travel"
    return "Other"


def parse_one_file(path: Path, log: list[str]) -> Optional[InvoiceResult]:
    """处理单个文件：识别类型并提取字段，不在这里移动文件。"""

    original_name = path.name
    suffix = path.suffix.lower()

    # PDF 优先走文本层，避免不必要 OCR；只有文本层为空或解析失败时才回退 Vision。
    pdf_text = ""
    if suffix == ".pdf":
        try:
            pdf_text = extract_pdf_text(path)
            if pdf_text.strip():
                train_result = parse_china_train_ticket(pdf_text, original_name)
                if train_result:
                    return train_result
                china_result = parse_china_standard_invoice(pdf_text, original_name)
                if china_result:
                    return china_result
                foreign_text_result = parse_foreign_program_text(pdf_text, original_name)
                if foreign_text_result:
                    return foreign_text_result
                log.append("  - PDF 有文本层，但未能按规则解析，将回退 Vision OCR。")
            else:
                log.append("  - PDF 文本层为空，将使用 Vision OCR。")
        except Exception as exc:
            log.append(f"  - PDF 文本层读取失败，将尝试 Vision OCR：{exc}")

    # 国外票据：按要求使用 macOS Vision OCR。
    lines = ocr_document(path)
    if not lines:
        raise RuntimeError("Vision OCR 未识别到文本")

    manual_result = parse_foreign_manual(lines, original_name)
    if manual_result:
        return manual_result

    program_result = parse_foreign_program(lines, original_name)
    if program_result:
        return program_result

    raise RuntimeError("未能匹配任何发票/票据类型")


def process_one_file(
    path: Path,
    log: list[str],
    seq: int = 1,
    mode: str = "Mainland",
    settings: Optional[AppSettings] = None,
    rates: Optional[RateManager] = None,
) -> Optional[InvoiceResult]:
    """兼容旧入口：解析后按模式换算、重命名、必要时标注。"""

    result = parse_one_file(path, log)
    if not result:
        return None
    return archive_file(path, result, seq=seq, mode=mode, settings=settings, rates=rates)


def currency_label(code: str) -> str:
    return "RMB" if normalize_currency_code(code) == "CNY" else normalize_currency_code(code)


def render_filename_stem(result: InvoiceResult, seq: int, settings: AppSettings) -> str:
    dt = datetime.strptime(result.date, "%Y-%m-%d")
    values = {
        "yyyy": f"{dt.year:04d}",
        "yy": f"{dt.year % 100:02d}",
        "mm": f"{dt.month:02d}",
        "dd": f"{dt.day:02d}",
        "yymm": f"{dt.year % 100:02d}{dt.month:02d}",
        "date": result.date,
        "seq": str(seq),
        "category": result.category,
        "currency": currency_label(result.report_currency),
        "amount": str(result.report_amount),
        "merchant": result.merchant,
        "type": result.invoice_type,
    }
    template = settings.filename_format or AppSettings.filename_format
    try:
        stem = template.format(**values)
    except Exception:
        stem = AppSettings.filename_format.format(**values)
    return sanitize_filename_part(stem, fallback=f"{values['yymm']}_{seq}_{result.category}_{values['currency']}{values['amount']}")


def archive_file(
    path: Path,
    result: InvoiceResult,
    annotate: Optional[bool] = None,
    seq: int = 1,
    mode: str = "Mainland",
    settings: Optional[AppSettings] = None,
    rates: Optional[RateManager] = None,
) -> InvoiceResult:
    settings = settings or AppSettings.load()
    rates = rates or RateManager()
    target_currency = "HKD" if mode.upper() == "HK" else "CNY"
    result.original_currency = normalize_currency_code(result.original_currency)
    result.report_currency = target_currency
    result.report_amount = round(rates.convert(result.amount_original, result.original_currency, target_currency))

    merchant = sanitize_filename_part(result.merchant)
    result.merchant = merchant
    result.category = sanitize_filename_part(result.category, fallback="Other")
    new_name = f"{render_filename_stem(result, seq, settings)}{path.suffix.lower()}"
    target = unique_path(path.with_name(new_name))

    archive_dir = path.parent / "归档"
    archive_dir.mkdir(exist_ok=True)
    archive_path = unique_path(archive_dir / path.name)
    shutil.move(str(path), str(archive_path))

    should_annotate = (
        settings.annotate_converted
        and result.invoice_type == "国外程序标注"
        and result.original_currency != result.report_currency
    )
    if annotate is not None:
        should_annotate = annotate

    if should_annotate:
        annotate_document_with_label(archive_path, f"{currency_label(result.report_currency)} {result.report_amount}", target)
    else:
        shutil.move(str(archive_path), str(target))

    result.new_name = target.name
    return result


def build_excel(results: list[InvoiceResult], folder: Path, report_currency: str = "CNY") -> Path:
    rows = []
    for item in results:
        dt = datetime.strptime(item.date, "%Y-%m-%d")
        rows.append(
            {
                "年份": f"{dt.year}",
                "月份": f"{dt.month:02d}",
                "日期": item.date,
                "Category": item.category,
                f"金额({report_currency})": item.report_amount,
                "原币种": item.original_currency,
                "原金额": round(item.amount_original, 2),
                "原文件名": item.original_name,
                "新文件名": item.new_name,
                "发票类型": item.invoice_type,
            }
        )

    excel_path = folder / "发票报销统计表.xlsx"
    amount_col = f"金额({report_currency})"
    df = pd.DataFrame(rows, columns=["年份", "月份", "日期", "Category", amount_col, "原币种", "原金额", "原文件名", "新文件名", "发票类型"])
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="报销统计")
        sheet = writer.book["报销统计"]
        widths = {
            "A": 10,
            "B": 10,
            "C": 14,
            "D": 12,
            "E": 14,
            "F": 10,
            "G": 12,
            "H": 34,
            "I": 42,
            "J": 18,
        }
        for col, width in widths.items():
            sheet.column_dimensions[col].width = width
    return excel_path


def iter_invoice_files(folder: Path) -> list[Path]:
    archive_dir = folder / "归档"
    return sorted(
        [
            path
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.lower() in SUPPORTED_EXTS
            and path.name != "发票报销统计表.xlsx"
            and path != archive_dir
        ]
    )


def normalize_folder_input(folder_input: str) -> Path:
    """兼容手动输入、终端拖拽生成的转义路径、以及 file:// 形式路径。"""

    folder_text = (folder_input or "").strip().strip("'\"")
    if folder_text.startswith("file://"):
        folder_text = urllib.parse.unquote(urllib.parse.urlparse(folder_text).path)
    # Finder 拖到终端常见格式：Family\ Shared/1.\ HK\ \&\ MC\ Company
    folder_text = folder_text.replace("\\ ", " ").replace("\\&", "&").replace("\\(", "(").replace("\\)", ")")
    return Path(folder_text).expanduser()
