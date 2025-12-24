# main.py
# -*- coding: utf-8 -*-
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi import Response
import traceback, logging
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import io
import os, json, tempfile, time, re, base64
from typing import List, Dict, Any, Optional, Tuple
from urllib.parse import urlparse, urljoin
from datetime import datetime
from collections import OrderedDict
from copy import deepcopy

import httpx
import trafilatura
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("aquilas")

app = FastAPI(title="Clipping Report Builder")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

UI_PATH = os.path.join(os.path.dirname(__file__), "ui.html")

@app.get("/", response_class=HTMLResponse)
def home():
    if os.path.exists(UI_PATH):
        return HTMLResponse(open(UI_PATH, "r", encoding="utf-8").read())
    return HTMLResponse("<h1>Clipping Report Builder</h1><p>Backend is running.</p>")

@app.head("/", include_in_schema=False)
def home_head():
    return Response(status_code=204)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------
# Clients
# -----------------------------
CLIENT_NAMES = [
    "ALPHA IVF GROUP BERHAD", "AME ELITE CONSORTIUM BERHAD", "AME REAL ESTATE INVESTMENT TRUST",
    "DELEUM BERHAD", "ECONPILE HOLDINGS BERHAD", "GAGASAN NADI CERGAS BERHAD",
    "GDB HOLDINGS BERHAD", "GUAN CHONG BERHAD", "KSL HOLDINGS BERHAD",
    "LAC MED BERHAD", "LAND & GENERAL BERHAD", "MALAYAN FLOUR MILLS BERHAD",
    "MALAYSIA STEEL WORKS (KL) BERHAD", "MATRIX CONCEPTS HOLDINGS BERHAD",
    "PECCA GROUP BERHAD", "SCIENTEX BERHAD", "SENHENG NEW RETAIL BERHAD",
    "SOUTHERN CABLE GROUP BERHAD", "TIONG NAM LOGISTICS HOLDINGS BERHAD", "TUJU SETIA BERHAD",
]

def _slug(s: str) -> str:
    out, dash = [], False
    for ch in s.lower():
        if ch.isalnum():
            out.append(ch)
            dash = False
        else:
            if not dash:
                out.append("-")
                dash = True
    return "".join(out).strip("-")

CLIENTS: List[Dict[str, str]] = [{"id": _slug(n), "name": n} for n in CLIENT_NAMES]

MEDIA_OPTIONS = [
    "Bernama", "The Star", "The Edge", "The Sun", "The Malaysian Reserve", "New Straits Times",
    "Sin Chew Daily", "China Press", "Nanyang Siang Pau", "Utusan Malaysia", "Berita Harian", "DagangNews",
]

# -----------------------------
# Models
# -----------------------------
class BuildItem(BaseModel):
    media: str
    category: str = "Online"
    section: Optional[str] = None
    language: Optional[str] = None
    date: Optional[str] = None
    url: Optional[str] = None
    snippet: Optional[str] = None
    title_override: Optional[str] = None
    image_data: Optional[str] = None

class BuildReportReq(BaseModel):
    client: str
    items: List[BuildItem]

# -----------------------------
# Helpers
# -----------------------------
def _resolve_client(client: str) -> Optional[Dict[str, str]]:
    key = (client or "").strip().lower()
    if not key: return None
    for c in CLIENTS:
        if key in {c["id"].lower(), c["name"].lower()}: return c
    return None

async def _fetch_html(url: str) -> str:
    async with httpx.AsyncClient(timeout=25, headers={"User-Agent": "Mozilla/5.0"}) as c:
        r = await c.get(url, follow_redirects=True)
        r.raise_for_status()
        return r.text

def _extract_article(url: str, html: str) -> Dict[str, Any]:
    raw = trafilatura.extract(html, url=url, include_images=False, include_links=False, output_format="json")
    if not raw: return {"title": "", "text": ""}
    try:
        j = json.loads(raw)
    except Exception:
        return {"title": "", "text": ""}
    return {"title": j.get("title") or "", "text": j.get("text") or ""}

IMG_META_PATTERNS = [
    re.compile(r'<meta\s+property=["\']og:image["\']\s+content=["\']([^"\']+)["\']', re.I),
    re.compile(r'<meta\s+name=["\']twitter:image["\']\s+content=["\']([^"\']+)["\']', re.I),
    re.compile(r'<link\s+rel=["\']image_src["\']\s+href=["\']([^"\']+)["\']', re.I),
]
IMG_FALLBACK_PAT = re.compile(r'<img\b[^>]*?\bsrc=["\']([^"\']+)["\'][^>]*>', re.I)

def _clone_para_format(src_p: Paragraph, dst_p: Paragraph):
    dst_p.style = src_p.style
    dst_p.alignment = src_p.alignment
    s, d = src_p.paragraph_format, dst_p.paragraph_format
    d.left_indent, d.right_indent = s.left_indent, s.right_indent
    d.first_line_indent, d.space_before, d.space_after = s.first_line_indent, s.space_before, s.space_after
    d.line_spacing, d.line_spacing_rule = s.line_spacing, s.line_spacing_rule

# ---------- Font helpers ----------
CJK_SEG_RE = re.compile(r'([\u4e00-\u9fff]+)')

def _segment_cjk(text: str):
    parts, last = [], 0
    text = text or ""
    for m in CJK_SEG_RE.finditer(text):
        s, e = m.span()
        if s > last: parts.append((text[last:s], False))
        parts.append((text[s:e], True))
        last = e
    if last < len(text): parts.append((text[last:], False))
    return parts

def _set_run_eastasia_font(run, ea_font: str = "SimSun", latin_font: Optional[str] = None):
    if latin_font: run.font.name = latin_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), ea_font)

def _rebuild_runs_cjk_aware(p: Paragraph, is_headline: bool = True, latin_font_name: str = "Trebuchet MS"):
    text_now = p.text or ""
    for r in reversed(p.runs):
        r._element.getparent().remove(r._element)
    for seg, is_cjk in _segment_cjk(text_now):
        if not seg: continue
        run = p.add_run(seg)
        run.bold = True  # <--- FIXED: Force bold for CJK segments
        if is_cjk:
            _set_run_eastasia_font(run, ea_font="SimSun")
        else:
            run.font.name = latin_font_name

# ---------- Article & Image ----------
def _normalize_article_paragraphs(text: str) -> list[str]:
    if not text: return []
    t = text.replace("\r\n", "\n").strip()
    if "\n\n" in t:
        return [re.sub(r'[ \t]+', ' ', b.strip()) for b in re.split(r'\n\s*\n', t) if b]
    return [re.sub(r'[ \t]+', ' ', ln.strip()) for ln in t.split("\n") if ln]

async def _download_image_to_temp(url: str) -> Optional[str]:
    try:
        async with httpx.AsyncClient(timeout=25) as c:
            r = await c.get(url, follow_redirects=True)
            r.raise_for_status()
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(r.content)
                return tmp.name
    except: return None

def _data_url_to_temp(data_url: str) -> Optional[str]:
    try:
        head, b64 = data_url.split(",", 1)
        raw = base64.b64decode(b64)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
            tmp.write(raw)
            return tmp.name
    except: return None

# ---------- Replacement Logic ----------
def _add_hyperlink(paragraph, url: str, text: str):
    part = paragraph.part
    r_id = part.relate_to(url, reltype=RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rPr.append(OxmlElement('w:b
