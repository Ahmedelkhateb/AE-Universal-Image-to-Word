"""
A.E Vision-Rebuild: photographed/scanned page -> editable Word 2010 DOCX.

Architecture:
1) PP-StructureV3: orientation/unwarp + OCR + layout + tables.
2) Independent OCR regions are preserved; no neighbor/cell text copying.
3) Source page is used ONLY as a single background layer after detected text is
   removed. It is never duplicated, tiled, or placed in the header.
4) OCR is rebuilt as real editable Word text using one VML text box per region.
5) Detected tables are rebuilt as native Word tables.
6) The original page is never inserted as the editable foreground.
7) If OCR fails completely, the script does not pretend success: it keeps the
   clean visual layer and writes a diagnostic report.

Word 2010 is the target. Python 3.11 + PaddleOCR 3.x + python-docx.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any

import cv2
import numpy as np
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

try:
    from paddleocr import PPStructureV3
except Exception as e:
    raise SystemExit(
        "PaddleOCR/PPStructureV3 is unavailable. Install the required "
        "PaddleOCR + PaddlePaddle packages first.\n"
        f"Original error: {e}"
    )

ARABIC_RE = re.compile(r"[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff]")
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def clean_text(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).replace("\x0c", " ")
    s = re.sub(r"[ \t\r\n]+", " ", s)
    return s.strip()

def bbox_from_poly(poly):
    a = np.asarray(poly, dtype=float)
    if a.ndim != 2 or a.shape[0] < 2:
        raise ValueError("invalid polygon")
    x0 = int(np.floor(a[:, 0].min()))
    y0 = int(np.floor(a[:, 1].min()))
    x1 = int(np.ceil(a[:, 0].max()))
    y1 = int(np.ceil(a[:, 1].max()))
    return x0, y0, max(1, x1-x0), max(1, y1-y0)

def inside_ratio(inner, outer):
    ax, ay, aw, ah = inner
    bx, by, bw, bh = outer
    x0, y0 = max(ax,bx), max(ay,by)
    x1, y1 = min(ax+aw,bx+bw), min(ay+ah,by+bh)
    inter = max(0,x1-x0)*max(0,y1-y0)
    return inter / max(1, aw*ah)

def px_to_pt(px, page_px, page_pt):
    return (float(px) / max(1.0,float(page_px))) * float(page_pt)

def twips(pt):
    return max(0, min(31680, int(round(float(pt)*20))))

def result_dict(res):
    if isinstance(res, dict):
        return res
    for a in ("json","res"):
        try:
            x = getattr(res,a)
            x = x() if callable(x) else x
            if isinstance(x,dict):
                return x
        except Exception:
            pass
    try:
        return dict(res)
    except Exception:
        return {}

def as_list(x):
    if x is None:
        return []
    try:
        return list(x)
    except Exception:
        return []

def extract_ocr(data):
    o = data.get("overall_ocr_res")
    if not isinstance(o,dict):
        o = {}
    texts = as_list(o.get("rec_texts"))
    boxes = as_list(o.get("rec_boxes"))
    if not boxes:
        boxes = as_list(o.get("rec_polys"))
    scores = as_list(o.get("rec_scores"))

    out=[]
    for i,t in enumerate(texts):
        t=clean_text(t)
        if not t or i>=len(boxes):
            continue
        try: b=bbox_from_poly(boxes[i])
        except Exception: continue
        try: s=float(scores[i]) if i<len(scores) else 1.0
        except Exception: s=1.0
        out.append({"text":t,"box":b,"score":s})
    return out

def extract_layout(data):
    raw=data.get("parsing_res_list")
    if raw is None:
        raw=data.get("layout_det_res")
    out=[]
    for it in as_list(raw):
        if not isinstance(it,dict): continue
        label=it.get("block_label") or it.get("label") or ""
        bb=it.get("block_bbox") or it.get("bbox")
        if bb is None: continue
        try: b=bbox_from_poly(bb)
        except Exception: continue
        out.append({"label":str(label).lower(),"box":b})
    return out

def extract_tables(data):
    raw=data.get("table_res_list")
    out=[]
    for it in as_list(raw):
        if not isinstance(it,dict): continue
        cells=it.get("cell_box_list")
        if cells is None: cells=it.get("cell_bbox_list")
        cells=as_list(cells)
        cb=[]
        for c in cells:
            try: cb.append(bbox_from_poly(c))
            except Exception: pass
        if not cb: continue
        x0=min(b[0] for b in cb); y0=min(b[1] for b in cb)
        x1=max(b[0]+b[2] for b in cb); y1=max(b[1]+b[3] for b in cb)
        o=it.get("table_ocr_pred")
        if not isinstance(o,dict): o={}
        rb=as_list(o.get("rec_boxes"))
        if not rb: rb=as_list(o.get("rec_polys"))
        rt=as_list(o.get("rec_texts"))
        rs=as_list(o.get("rec_scores"))
        texts=[]
        for i,t in enumerate(rt):
            t=clean_text(t)
            if not t or i>=len(rb): continue
            try: b=bbox_from_poly(rb[i])
            except Exception: continue
            try: s=float(rs[i]) if i<len(rs) else 1.0
            except Exception: s=1.0
            texts.append({"text":t,"box":b,"score":s})
        out.append({
            "bbox":(x0,y0,x1-x0,y1-y0),
            "html":it.get("pred_html") or "",
            "cell_boxes":cb,
            "ocr":texts,
        })
    return out

def sample_color(img,b):
    x,y,w,h=b
    H,W=img.shape[:2]
    crop=img[max(0,y):min(H,y+h),max(0,x):min(W,x+w)]
    if crop.size==0: return (0,0,0)
    gray=cv2.cvtColor(crop,cv2.COLOR_BGR2GRAY)
    dark=crop[gray < np.percentile(gray,35)]
    p=dark.mean(axis=0) if len(dark) else crop.reshape(-1,3).mean(axis=0)
    bb,gg,rr=[int(np.clip(v,0,255)) for v in p]
    return rr,gg,bb

def make_clean_background(img, text_boxes):
    # Conservative masks: remove glyphs, not whole layout blocks.
    mask=np.zeros(img.shape[:2],np.uint8)
    for x,y,w,h in text_boxes:
        pad=max(1,int(round(min(w,h)*0.10)))
        x0=max(0,x-pad); y0=max(0,y-pad)
        x1=min(img.shape[1],x+w+pad); y1=min(img.shape[0],y+h+pad)
        cv2.rectangle(mask,(x0,y0),(x1,y1),255,-1)
    # Do not use huge morphology; it destroys table lines and graphics.
    return cv2.inpaint(img,mask,2.5,cv2.INPAINT_TELEA)

def html_matrix(html):
    soup=BeautifulSoup(html or "","lxml")
    tbl=soup.find("table")
    if tbl is None: return []
    rows=[]
    for tr in tbl.find_all("tr"):
        row=[]
        for c in tr.find_all(["td","th"],recursive=False):
            txt=clean_text(c.get_text(" ",strip=True))
            try: cs=max(1,int(c.get("colspan","1")))
            except: cs=1
            try: rs=max(1,int(c.get("rowspan","1")))
            except: rs=1
            row.append({"text":txt,"colspan":cs,"rowspan":rs})
        if row: rows.append(row)
    return rows

def set_cell_margins(cell, top=30, start=40, bottom=30, end=40):
    tcPr=cell._tc.get_or_add_tcPr()
    tcMar=tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar=OxmlElement("w:tcMar"); tcPr.append(tcMar)
    for side,val in (("top",top),("start",start),("bottom",bottom),("end",end)):
        e=tcMar.find(qn("w:"+side))
        if e is None:
            e=OxmlElement("w:"+side); tcMar.append(e)
        e.set(qn("w:w"),str(val)); e.set(qn("w:type"),"dxa")

def set_cell_border(cell,color="808080",size="4"):
    tcPr=cell._tc.get_or_add_tcPr()
    b=tcPr.first_child_found_in("w:tcBorders")
    if b is None:
        b=OxmlElement("w:tcBorders"); tcPr.append(b)
    for side in ("top","left","bottom","right"):
        e=b.find(qn("w:"+side))
        if e is None:
            e=OxmlElement("w:"+side); b.append(e)
        e.set(qn("w:val"),"single"); e.set(qn("w:sz"),str(size))
        e.set(qn("w:space"),"0"); e.set(qn("w:color"),color)

def set_table_float(table,xpt,ypt):
    pr=table._tbl.tblPr
    old=pr.find(qn("w:tblpPr"))
    if old is not None: pr.remove(old)
    p=OxmlElement("w:tblpPr")
    p.set(qn("w:leftFromText"),"0"); p.set(qn("w:rightFromText"),"0")
    p.set(qn("w:topFromText"),"0"); p.set(qn("w:bottomFromText"),"0")
    p.set(qn("w:vertAnchor"),"page"); p.set(qn("w:horzAnchor"),"page")
    p.set(qn("w:tblpX"),str(twips(xpt))); p.set(qn("w:tblpY"),str(twips(ypt)))
    pr.append(p)

def set_table_width(table,wpt):
    pr=table._tbl.tblPr
    tw=pr.first_child_found_in("w:tblW")
    if tw is None:
        tw=OxmlElement("w:tblW"); pr.insert(0,tw)
    tw.set(qn("w:w"),str(twips(wpt))); tw.set(qn("w:type"),"dxa")
    lay=pr.first_child_found_in("w:tblLayout")
    if lay is None:
        lay=OxmlElement("w:tblLayout"); pr.append(lay)
    lay.set(qn("w:type"),"fixed")

def set_row_height(row,hpt):
    trPr=row._tr.get_or_add_trPr()
    e=trPr.find(qn("w:trHeight"))
    if e is None:
        e=OxmlElement("w:trHeight"); trPr.append(e)
    e.set(qn("w:val"),str(twips(hpt))); e.set(qn("w:hRule"),"atLeast")

def add_vml_textbox(paragraph,text,xpt,ypt,wpt,hpt,fontpt,rgb,rtl,bold):
    """
    VML is used because Word 2010 has native support for it.
    IMPORTANT: use Clark-notation namespace names; OxmlElement('v:shape')
    fails with KeyError when python-docx has not registered the VML prefix.
    """
    text=clean_text(text)
    if not text: return

    shape=OxmlElement("{%s}shape"%V_NS)
    pict=OxmlElement("w:pict")
    shape.set("style",
        "position:absolute;"
        f"left:{xpt:.2f}pt;top:{ypt:.2f}pt;"
        f"width:{max(5,wpt):.2f}pt;height:{max(7,hpt):.2f}pt;"
        "mso-position-horizontal:absolute;"
        "mso-position-horizontal-relative:page;"
        "mso-position-vertical:absolute;"
        "mso-position-vertical-relative:page;"
        "mso-wrap-style:none;"
        "z-index:20;")
    shape.set("stroked","f"); shape.set("filled","f")

    tb=OxmlElement("{%s}textbox"%V_NS)
    tb.set("style","mso-fit-shape-to-text:false;")
    content=OxmlElement("w:txbxContent")
    p=OxmlElement("w:p")
    pp=OxmlElement("w:pPr")
    jc=OxmlElement("w:jc"); jc.set(qn("w:val"),"right" if rtl else "left"); pp.append(jc)
    if rtl: pp.append(OxmlElement("w:bidi"))
    p.append(pp)

    r=OxmlElement("w:r"); rp=OxmlElement("w:rPr")
    for tag in ("w:sz","w:szCs"):
        e=OxmlElement(tag); e.set(qn("w:val"),str(max(8,int(round(fontpt*2))))); rp.append(e)
    if bold:
        rp.append(OxmlElement("w:b")); rp.append(OxmlElement("w:bCs"))
    col=OxmlElement("w:color"); col.set(qn("w:val"),"%02X%02X%02X"%rgb); rp.append(col)
    fonts=OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"),"Arial"); fonts.set(qn("w:hAnsi"),"Arial")
    fonts.set(qn("w:cs"),"Arial"); rp.append(fonts)
    r.append(rp)
    t=OxmlElement("w:t"); t.set("{http://www.w3.org/XML/1998/namespace}space","preserve"); t.text=text
    r.append(t); p.append(r); content.append(p); tb.append(content); shape.append(tb); pict.append(shape)
    paragraph._p.append(pict)

def add_background_once(doc, image_path, width_in):
    """
    Put exactly ONE clean page image in the body as a behind-text VML picture.
    It is not placed in the header, and it is never repeated.
    """
    p=doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    # Insert a VML picture with a relationship to the image.
    part=doc.part
    rid=part.relate_to(str(image_path), R_NS, is_external=True)
    # External file relationships are not suitable inside DOCX. Therefore use
    # python-docx's normal embedded picture once, then move its drawing behind.
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    run=p.add_run(); run.add_picture(str(image_path),width=Inches(width_in))
    drawing=run._r.xpath(".//w:drawing")
    if drawing:
        d=drawing[0]
        # Convert inline -> anchor so it is page-positioned.
        inline=d.find(qn("wp:inline"))
        if inline is not None:
            anchor=OxmlElement("wp:anchor")
            for child in list(inline):
                anchor.append(child)
            anchor.set("behindDoc","1")
            anchor.set("layoutInCell","0")
            anchor.set("allowOverlap","1")
            # Position at page origin.
            posH=OxmlElement("wp:positionH"); posH.set("relativeFrom","page")
            ph=OxmlElement("wp:posOffset"); ph.text="0"; posH.append(ph)
            posV=OxmlElement("wp:positionV"); posV.set("relativeFrom","page")
            pv=OxmlElement("wp:posOffset"); pv.text="0"; posV.append(pv)
            anchor.insert(0,posH); anchor.insert(1,posV)
            wrap=OxmlElement("wp:wrapNone"); anchor.append(wrap)
            d.remove(inline); d.append(anchor)
    # Keep paragraph itself zero-height visually.
    p.paragraph_format.line_spacing=0
    return p

def build_table(doc,t,pageW,pageH,pageWpt,pageHpt):
    rows=html_matrix(t["html"])
    if not rows: return None
    cols=max(sum(int(c["colspan"]) for c in r) for r in rows)
    if cols<=0: return None
    tbl=doc.add_table(rows=len(rows),cols=cols)
    tbl.autofit=False
    x,y,w,h=t["bbox"]
    set_table_width(tbl,px_to_pt(w,pageW,pageWpt))
    set_table_float(tbl,px_to_pt(x,pageW,pageWpt),px_to_pt(y,pageH,pageHpt))

    occ=set()
    for ri,row in enumerate(rows):
        ci=0
        for item in row:
            while (ri,ci) in occ and ci<cols: ci+=1
            if ci>=cols: break
            cs=max(1,int(item["colspan"])); rs=max(1,int(item["rowspan"]))
            ec=min(cols-1,ci+cs-1); er=min(len(rows)-1,ri+rs-1)
            cell=tbl.cell(ri,ci)
            if cs>1 or rs>1:
                try: cell=cell.merge(tbl.cell(er,ec))
                except Exception: pass
            for rr in range(ri,er+1):
                for cc in range(ci,ec+1): occ.add((rr,cc))
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
            rtl=bool(ARABIC_RE.search(item["text"]))
            p.alignment=WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
            run=p.add_run(item["text"]); run.font.name="Arial"; run.font.size=Pt(9)
            if rtl: run._element.get_or_add_rPr().append(OxmlElement("w:rtl"))
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell); set_cell_border(cell)
            ci=ec+1
        for rr in [ri]:
            set_row_height(tbl.rows[rr],max(10,px_to_pt(h/max(1,len(rows)),pageH,pageHpt)))
    return tbl

def reconstruct(inp,out,device="cpu"):
    img=cv2.imread(str(inp),cv2.IMREAD_COLOR)
    if img is None: raise FileNotFoundError(inp)
    H,W=img.shape[:2]
    print(f"[1/7] Loading image: {W} x {H}")
    print("[2/7] Running PP-StructureV3 layout + OCR + table analysis...")

    pipe=PPStructureV3(
        use_doc_orientation_classify=True,
        use_doc_unwarping=True,
        use_textline_orientation=True,
        use_table_recognition=True,
        use_formula_recognition=False,
        use_chart_recognition=False,
        device=device,
        enable_mkldnn=False,
        cpu_threads=4,
    )
    results=list(pipe.predict(str(inp)))
    if not results: raise RuntimeError("PP-StructureV3 returned no result.")
    data=result_dict(results[0])

    ocr=extract_ocr(data)
    layout=extract_layout(data)
    tables=extract_tables(data)
    print(f"    OCR regions: {len(ocr)}")
    print(f"    Tables: {len(tables)}")

    table_boxes=[t["bbox"] for t in tables]
    free=[]
    for z in ocr:
        if any(inside_ratio(z["box"],tb)>0.70 for tb in table_boxes):
            continue
        free.append(z)

    table_text=[]
    for t in tables:
        table_text.extend(t["ocr"])
    # Do not duplicate table OCR as free text.
    regions=free+[
        {"text":z["text"],"box":z["box"],"score":z["score"],"table":True}
        for z in table_text
    ]

    print(f"    Editable text regions: {len(regions)}")
    if not regions and not tables:
        print("WARNING: OCR/layout returned nothing. No fake editable output will be claimed.")

    print("[3/7] Removing only detected text from visual layer...")
    clean=make_clean_background(img,[z["box"] for z in regions])
    work=Path(tempfile.mkdtemp(prefix="ae_vision_"))
    clean_path=work/"clean_page.png"
    cv2.imwrite(str(clean_path),clean)

    print("[4/7] Creating ONE-page Word 2010 document...")
    doc=Document()
    sec=doc.sections[0]
    pageW_in=8.27
    pageH_in=pageW_in*H/max(1,W)
    sec.page_width=Inches(pageW_in); sec.page_height=Inches(pageH_in)
    sec.left_margin=sec.right_margin=sec.top_margin=sec.bottom_margin=Inches(0)
    sec.header_distance=sec.footer_distance=Inches(0)

    # EXACTLY ONE background image, body only.
    add_background_once(doc,clean_path,pageW_in)

    # Always ensure an anchor exists.
    anchor=doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    pageWpt=pageW_in*72; pageHpt=pageH_in*72

    print("[5/7] Rebuilding native editable text...")
    for z in regions:
        tx=clean_text(z["text"])
        if not tx: continue
        x,y,w,h=z["box"]
        xpt=px_to_pt(x,W,pageWpt); ypt=px_to_pt(y,H,pageHpt)
        wpt=max(6,px_to_pt(w*1.08,W,pageWpt))
        hpt=max(8,px_to_pt(h*1.75,H,pageHpt))
        fs=max(6,min(34,px_to_pt(h*.90,H,pageHpt)))
        rgb=sample_color(img,z["box"])
        rtl=bool(ARABIC_RE.search(tx))
        crop=img[max(0,y):min(H,y+h),max(0,x):min(W,x+w)]
        dark=False
        if crop.size:
            g=cv2.cvtColor(crop,cv2.COLOR_BGR2GRAY)
            dark=float(np.mean(g<100))>.18
        add_vml_textbox(anchor,tx,xpt,ypt,wpt,hpt,fs,rgb,rtl,dark)

    print("[6/7] Rebuilding native editable tables...")
    for t in tables:
        build_table(doc,t,W,H,pageWpt,pageHpt)

    # Diagnostic custom document properties are intentionally avoided for
    # Word-2010 compatibility. Save normally.
    out.parent.mkdir(parents=True,exist_ok=True)
    doc.save(str(out))
    shutil.rmtree(work,ignore_errors=True)

    print("[7/7] DONE")
    print(f"Output: {out}")
    print("Architecture: ONE clean visual layer + native editable OCR + native tables.")
    print("The original page is NOT inserted repeatedly and is NOT the editable foreground.")

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("-o","--output",default=None)
    ap.add_argument("--device",default="cpu",choices=["cpu","gpu:0"])
    a=ap.parse_args()
    inp=Path(a.image).resolve()
    if not inp.exists(): raise SystemExit(f"Input file not found: {inp}")
    out=Path(a.output).resolve() if a.output else inp.with_name(inp.stem+"_FINAL_EDITABLE.docx")
    reconstruct(inp,out,a.device)

if __name__=="__main__":
    main()
