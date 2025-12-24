# main.py
# -*- coding: utf-8 -*-
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi import Response
import traceback, logging
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import os, json, tempfile, re, base64
from typing import List, Dict, Any, Optional, Tuple
from urllib.parse import urlparse, urljoin
from datetime import datetime
from collections import OrderedDict
from copy import deepcopy

import httpx
import trafilatura
from docx import Document
from docx.shared import Pt, Cm
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
MEDIA_OPTIONS = ["Bernama", "The Star", "The Edge", "The Sun", "The Malaysian Reserve", "New Straits Times", "Sin Chew Daily", "China Press", "Nanyang Siang Pau", "Utusan Malaysia", "Berita Harian", "DagangNews"]

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
# Bold & Font Helpers
# -----------------------------
def _set_run_bold(run):
    """Force bold on a run by setting property to True and adding XML b tag."""
    run.bold = True
    rPr = run._element.get_or_add_rPr()
    b = rPr.find(qn('w:b'))
    if b is None:
        b = OxmlElement('w:b')
        rPr.append(b)
    b.set(qn('w:val'), '1')

def _set_run_eastasia_font(run, ea_font="SimSun"):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), ea_font)

def _rebuild_runs_cjk_aware(p: Paragraph, latin_font_name="Trebuchet MS"):
    text_now = p.text or ""
    for r in list(p.runs)[::-1]:
        r._element.getparent().remove(r._element)
    
    parts = []
    last = 0
    for m in re.finditer(r'([\u4e00-\u9fff]+)', text_now):
        s, e = m.span()
        if s > last: parts.append((text_now[last:s], False))
        parts.append((text_now[s:e], True))
        last = e
    if last < len(text_now): parts.append((text_now[last:], False))

    for seg, is_cjk in parts:
        if not seg: continue
        run = p.add_run(seg)
        _set_run_bold(run) # Force Bold
        if is_cjk:
            _set_run_eastasia_font(run, "SimSun")
        else:
            run.font.name = latin_font_name

# -----------------------------
# Core Replacement Logic
# -----------------------------
def _add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, reltype=RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rPr.append(OxmlElement('w:b')) # Hyperlink Bold
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rPr.append(c)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.text = text; new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def _replace_in_paragraph_single(p: Paragraph, mapping: Dict[str, Any], url_for_link: Optional[str]):
    original_text = p.text or ""
    had_url = "{{ITEM_URL}}" in original_text
    
    # Text Replacement
    final_text = original_text
    for k, v in mapping.items():
        if k in ("ITEM_IMAGE", "ITEM_URL"): continue
        final_text = final_text.replace(f"{{{{{k}}}}}", str(v or ""))

    # Clear and rewrite with BOLD
    if not had_url:
        for r in list(p.runs)[::-1]:
            r._element.getparent().remove(r._element)
        new_run = p.add_run(final_text)
        _set_run_bold(new_run) # Force Bold
    else:
        # Handle URL special line
        for r in list(p.runs)[::-1]:
            r._element.getparent().remove(r._element)
        if url_for_link:
            prefix = p.add_run("Source from: ")
            _set_run_bold(prefix)
            _add_hyperlink(p, url_for_link, url_for_link)

def _replace_placeholders_in_inserted_elements(doc, elements, mapping, img_path, url, use_cjk, is_newspaper):
    found_img = False
    for el in elements:
        if isinstance(el, CT_P):
            p = Paragraph(el, doc)
            txt = p.text or ""
            if "{{ITEM_IMAGE}}" in txt:
                found_img = True
                _insert_image_into_paragraph(p, img_path, is_newspaper)
            else:
                _replace_in_paragraph_single(p, mapping, url)
                if use_cjk: _rebuild_runs_cjk_aware(p)
                else:
                    for r in p.runs: _set_run_bold(r) # Final Force Bold

        elif isinstance(el, CT_Tbl):
            tbl = Table(el, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        txt = p.text or ""
                        if "{{ITEM_IMAGE}}" in txt:
                            found_img = True
                            _insert_image_into_paragraph(p, img_path, is_newspaper)
                        else:
                            _replace_in_paragraph_single(p, mapping, url)
                            if use_cjk: _rebuild_runs_cjk_aware(p)
                            else:
                                for r in p.runs: _set_run_bold(r) # Force Bold inside Table

# -----------------------------
# Scrapers & Utils
# -----------------------------
def _is_online(cat): return "online" in (cat or "").lower()
def _fmt_date(s):
    try: return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%d %B %Y")
    except: return s or ""

async def _fetch_html(url):
    async with httpx.AsyncClient(timeout=25, headers={"User-Agent": "Mozilla/5.0"}) as c:
        r = await c.get(url, follow_redirects=True)
        r.raise_for_status()
        return r.text

async def _fetch_article_fields(item: BuildItem):
    res = {"title": item.title_override or "[Headline]", "text": item.snippet or "", "image_url": None, "pasted_image": None}
    if item.image_data: res["pasted_image"] = _data_url_to_temp(item.image_data)
    if _is_online(item.category) and item.url:
        try:
