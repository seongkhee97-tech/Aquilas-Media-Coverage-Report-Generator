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
    # Loop through each run in the paragraph to find and replace placeholders
    # This preserves the formatting (like Bold) of individual runs
    for run in p.runs:
        original_run_text = run.text
        new_run_text = original_run_text
        
        for k, v in mapping.items():
            placeholder = f"{{{{{k}}}}}"
            if placeholder in new_run_text:
                new_run_text = new_run_text.replace(placeholder, str(v or ""))
        
        if new_run_text != original_run_text:
            run.text = new_run_text
            # If you want the replaced value to ALSO be bold, keep this:
            run.bold = True

    # Special handling for URLs since they often require clearing the whole line
    if "{{ITEM_URL}}" in p.text:
        for r in list(p.runs)[::-1]:
            r._element.getparent().remove(r._element)
        if url_for_link:
            # Re-add bold prefix to match your template style
            prefix = p.add_run("Source from: ")
            prefix.bold = True
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
            h = await _fetch_html(item.url)
            ex = trafilatura.extract(h, url=item.url, include_images=False, output_format="json")
            if ex:
                j = json.loads(ex)
                if not item.title_override: res["title"] = j.get("title") or res["title"]
                if not item.snippet: res["text"] = j.get("text") or res["text"]
        except: pass
    return res

def _data_url_to_temp(d):
    try:
        raw = base64.b64decode(d.split(",")[1])
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as t:
            t.write(raw); return t.name
    except: return None

async def _download_image_to_temp(u):
    try:
        async with httpx.AsyncClient() as c:
            r = await c.get(u); r.raise_for_status()
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as t:
                t.write(r.content); return t.name
    except: return None

def _insert_image_into_paragraph(p, path, keep_orig):
    for r in list(p.runs)[::-1]: r._element.getparent().remove(r._element)
    if path:
        try:
            r = p.add_run()
            if keep_orig: r.add_picture(path)
            else: r.add_picture(path, width=Cm(14))
        except: pass

# -----------------------------
# Build logic
# -----------------------------
async def _build_with_template(t_path, req, client_name):
    doc = Document(t_path)
    blocks = list(doc.element.body.iterchildren())
    
    # 1. Page-level summary replacement
    mapping_doc = {"CLIENT_NAME": client_name, "TOTAL_CLIPPINGS": len(req.items)}
    for p in doc.paragraphs:
        for k, v in mapping_doc.items():
            if f"{{{{{k}}}}}" in p.text: p.text = p.text.replace(f"{{{{{k}}}}}", str(v))

    # 2. Extract Card Proto
    marker_idx = -1
    card_proto = []
    for i, b in enumerate(blocks):
        if isinstance(b, CT_P) and "{{EXTRACTS_START}}" in (Paragraph(b, doc).text or ""):
            marker_idx = i
            j = i + 1
            while j < len(blocks):
                txt = Paragraph(blocks[j], doc).text if isinstance(blocks[j], CT_P) else "TABLE"
                if any(tok in txt for tok in ["{{ITEM_HEADLINE}}", "{{ITEM_MEDIA}}", "TABLE"]):
                    card_proto.append(deepcopy(blocks[j]))
                    j += 1
                else: break
            break
    
    # Clean up marker and proto from document
    if marker_idx != -1:
        marker_p = Paragraph(blocks[marker_idx], doc)
        marker_p.text = ""
        ref_el = blocks[marker_idx]
        for i in range(len(card_proto)):
            blocks[marker_idx + 1 + i].getparent().remove(blocks[marker_idx + 1 + i])

        # 3. Generate Items
        for it in req.items:
            art = await _fetch_article_fields(it)
            is_news = not _is_online(it.category)
            mapping = {
                "ITEM_HEADLINE": art["title"] if not is_news else "",
                "ITEM_CONTENT": art["text"] if not is_news else "",
                "ITEM_MEDIA": it.media or "", "ITEM_SECTION": it.section or "",
                "ITEM_LANGUAGE": it.language or "", "ITEM_DATE": _fmt_date(it.date)
            }
            img = art["pasted_image"] or (await _download_image_to_temp(art["image_url"]) if art["image_url"] else None)
            
            inserted = []
            for xml_el in card_proto:
                new_el = deepcopy(xml_el)
                ref_el.addnext(new_el)
                ref_el = new_el
                inserted.append(new_el)
            
            _replace_placeholders_in_inserted_elements(doc, inserted, mapping, img, it.url, it.language=="Chinese", is_news)
            
    return doc

@app.post("/build_report")
async def build_report(req: BuildReportReq):
    try:
        c = next((cl for cl in CLIENTS if cl["id"] == req.client), {"name": req.client})
        t_path = os.path.join(os.path.dirname(__file__), "template.docx")
        doc = await _build_with_template(t_path, req, c["name"])
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, "rb") as f: data = f.read()
        os.remove(tmp.name)
        
        return Response(content=data, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        headers={"Content-Disposition": f"attachment; filename=report.docx"})
    except Exception as e:
        logger.exception("Build failed")
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/clients")
def get_clients(): return CLIENTS
@app.get("/media")
def get_media(): return MEDIA_OPTIONS

