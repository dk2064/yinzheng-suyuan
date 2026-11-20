#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_page.py — 给下载到的文献正文前插入一页「审核页」。

两个子命令：
  build  : 依据 audit_meta.json 生成审核页 docx（每个条目一份）
  merge  : 把 docx 导出的审核页 PDF 合并到正文 PDF 前面，输出 *_审核版.pdf 与 audit_index.tsv

审核页内容：题名 / 作者 / 刊名 / 年卷期页 / 收录标识 / 核验结论与修正说明 /
GB/T 7714-2015 与 APA 引用串 / 文中引用写法 / 文章链接 / 文件指纹（SHA-256）。

用法（配合 add_audit_page.ps1，Word 负责 docx→pdf）：
  python audit_page.py build --meta audit_meta.json --docx-dir _covers
  python audit_page.py merge --meta audit_meta.json --cover-dir _covers --out fulltext_audited
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

CJK_FONT = "微软雅黑"
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x5A, 0x5A)
ACCENT = "1F3A5F"
OK_COLOR = "1E7B34"
WARN_COLOR = "B8791A"
BAD_COLOR = "A32B2B"

VERDICT = {
    "confirmed": ("通过", OK_COLOR, "题名、作者、刊名、年卷期页与原文一致"),
    "corrected": ("通过 · 需修正", WARN_COLOR, "条目真实存在，但元数据有误，引用前按下方说明改正"),
    "unverified": ("未能核实", BAD_COLOR, "公开源未检出对应记录，不建议直接引用"),
}


# ---------- docx helpers ----------
def set_run_font(run, size=10.5, bold=False, color=INK, font=CJK_FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)


def shade(element, hex_fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    element.append(shd)


def para_shade(p, hex_fill):
    shade(p._p.get_or_add_pPr(), hex_fill)


def cell_shade(cell, hex_fill):
    shade(cell._tc.get_or_add_tcPr(), hex_fill)


def add_hyperlink(paragraph, url, text, size=9.5):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), CJK_FONT)
    rFonts.set(qn("w:hAnsi"), CJK_FONT)
    rFonts.set(qn("w:eastAsia"), CJK_FONT)
    rPr.append(rFonts)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1155CC")
    rPr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def tight(p, before=0, after=2, line=1.15):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    return p


# ---------- citation strings ----------
def author_str(authors, sep=", "):
    if not authors:
        return ""
    if len(authors) <= 3:
        return sep.join(authors)
    return sep.join(authors[:3]) + ", 等"


def cite_gbt(it):
    a = author_str(it["authors"])
    t = it["title"]
    if it["kind"] == "学位论文":
        return f"{a}. {t}[D]. {it.get('city','')}: {it['venue']}, {it['year']}."
    head = f"{a}. {t}[J]. {it['venue']}, {it['year']}"
    if it.get("volume"):
        head += f", {it['volume']}"
    if it.get("issue"):
        head += f"({it['issue']})"
    if it.get("pages"):
        head += f": {it['pages']}"
    return head + "."


def cite_apa(it):
    a = author_str(it["authors"])
    t = it["title"]
    ids = it.get("ids") or {}
    if it["kind"] == "学位论文":
        s = f"{a}. ({it['year']}). {t} [硕士学位论文]. {it['venue']}."
    else:
        s = f"{a}. ({it['year']}). {t}. {it['venue']}"
        if it.get("volume"):
            s += f", {it['volume']}"
        if it.get("issue"):
            s += f"({it['issue']})"
        if it.get("pages"):
            s += f", {it['pages']}"
        s += "."
    if ids.get("doi"):
        s += f" https://doi.org/{ids['doi']}"
    return s


def cite_inline(it):
    a = it["authors"]
    if len(a) == 1:
        name = a[0]
    elif len(a) == 2:
        name = f"{a[0]} 和 {a[1]}"
    else:
        name = a[0] + " 等"
    return f"（{name}，{it['year']}）"


def id_str(it):
    ids = it.get("ids") or {}
    parts = []
    if ids.get("doi"):
        parts.append(f"DOI {ids['doi']}")
    if ids.get("cnki"):
        parts.append(f"知网 {ids['cnki']}")
    if ids.get("wanfang"):
        parts.append(f"万方 {ids['wanfang']}")
    if ids.get("cnki_cdmd"):
        parts.append(f"知网学位 {ids['cnki_cdmd']}")
    if ids.get("article_no"):
        parts.append(f"文章编号 {ids['article_no']}")
    if ids.get("issn"):
        parts.append(f"ISSN {ids['issn']}")
    return "；".join(parts) if parts else "—"


def vol_issue_pages(it):
    if it["kind"] == "学位论文":
        return f"{it['year']} 年 · 硕士学位论文"
    s = it["year"]
    if it.get("volume"):
        s += f", {it['volume']}"
    if it.get("issue"):
        s += f"({it['issue']})"
    if it.get("pages"):
        s += f": {it['pages']}"
    return s


# ---------- audit page ----------
def build_docx(it, batch, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for m in ("left_margin", "right_margin"):
        setattr(sec, m, Cm(1.9))
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.4)

    style = doc.styles["Normal"]
    style.font.name = CJK_FONT
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), CJK_FONT)
    style.paragraph_format.space_after = Pt(0)

    # 顶部色带（左右分栏，右侧编号贴右边界）
    p = doc.add_paragraph()
    para_shade(p, ACCENT)
    tight(p, 0, 0).paragraph_format.left_indent = Cm(0.15)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.9), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(
        f"  文献审核页  ·  LITERATURE VERIFICATION SHEET\tNo.{it['no']:02d} / {batch['total']}  "
    )
    set_run_font(r, 9.5, True, RGBColor(0xFF, 0xFF, 0xFF))

    # 题名
    p = doc.add_paragraph()
    tight(p, 10, 2)
    r = p.add_run(it["title"])
    set_run_font(r, 15, True, INK)

    # 基本信息表
    rows = [
        ("作　　者", "、".join(it["authors"])),
        ("文献类型", it["kind"] + ("（硕士）" if it["kind"] == "学位论文" else "")),
        ("来源载体", it["venue"]),
        ("年卷期页", vol_issue_pages(it)),
        ("收录标识", id_str(it)),
    ]
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, (k, v) in enumerate(rows):
        c0, c1 = tbl.rows[i].cells
        c0.width, c1.width = Cm(3.2), Cm(13.9)
        cell_shade(c0, "F2F4F7")
        p0 = c0.paragraphs[0]
        tight(p0)
        set_run_font(p0.add_run(k), 9.5, True, MUTED)
        p1 = c1.paragraphs[0]
        tight(p1)
        set_run_font(p1.add_run(v or "—"), 9.5, False, INK)

    # 核验结论
    label, color, gloss = VERDICT.get(it["verdict"], VERDICT["unverified"])
    p = doc.add_paragraph()
    tight(p, 10, 0)
    set_run_font(p.add_run("核验结论　"), 10.5, True, MUTED)
    set_run_font(p.add_run("✔ " + label), 11, True, RGBColor.from_string(color))
    set_run_font(p.add_run(f"　{gloss}"), 9.5, False, MUTED)
    if it.get("notes"):
        p = doc.add_paragraph()
        tight(p, 2, 0)
        set_run_font(p.add_run(it["notes"]), 9.5, False, INK)

    # 引用格式
    def head(text):
        p = doc.add_paragraph()
        tight(p, 12, 3)
        set_run_font(p.add_run(text), 10.5, True, RGBColor.from_string(ACCENT))

    head("引用格式")
    p = doc.add_paragraph()
    tight(p)
    set_run_font(p.add_run("GB/T 7714-2015　"), 9.5, True, MUTED)
    set_run_font(p.add_run(cite_gbt(it)), 9.5, False, INK)
    p = doc.add_paragraph()
    tight(p)
    set_run_font(p.add_run("APA（第 7 版）　"), 9.5, True, MUTED)
    set_run_font(p.add_run(cite_apa(it)), 9.5, False, INK)
    p = doc.add_paragraph()
    tight(p)
    set_run_font(p.add_run("文中引用　"), 9.5, True, MUTED)
    set_run_font(p.add_run(cite_inline(it)), 9.5, False, INK)

    # 文章链接
    head("文章链接")
    for lk in it.get("links") or []:
        p = doc.add_paragraph()
        tight(p, 1, 1)
        set_run_font(p.add_run(f"{lk['label']}　"), 9.5, True, MUTED)
        add_hyperlink(p, lk["url"], lk["url"], size=7.8)
    if not it.get("links"):
        p = doc.add_paragraph()
        tight(p)
        set_run_font(p.add_run("—"), 9.5, False, MUTED)

    # 文件信息
    head("文件与来源")
    loc = it.get("local") or {}
    fname = loc.get("out_name") or loc.get("file") or "—"
    info = [
        ("文件名", fname),
        ("正文页数", str(loc.get("pages", "—")) if loc.get("pages") else "—"),
        ("下载渠道", loc.get("source") or "—"),
        ("文件指纹", (loc.get("sha256", "")[:32] + "…") if loc.get("sha256") else "—"),
    ]
    for k, v in info:
        p = doc.add_paragraph()
        tight(p, 1, 1)
        set_run_font(p.add_run(f"{k}　"), 9.5, True, MUTED)
        set_run_font(p.add_run(v), 8.5, False, INK)

    # 页脚
    p = doc.add_paragraph()
    tight(p, 14, 0)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    bar = doc.add_paragraph()
    para_shade(bar, "E8ECF2")
    tight(bar, 0, 0)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    set_run_font(
        bar.add_run(f"  审核时间 {stamp}　·　来源 {batch['source']}　·　生成工具 yinzheng-suyuan"),
        8,
        False,
        MUTED,
    )
    p = doc.add_paragraph()
    tight(p, 3, 0)
    set_run_font(p.add_run("正文自下一页开始。"), 8.5, False, MUTED)

    doc.save(path)


def sha256_of(path, limit=None):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(1 << 20)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def slug(s, n=24):
    s = re.sub(r"[\\/:*?\"<>|\s]+", "", s)
    return s[:n]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["build", "merge"])
    ap.add_argument("--meta", required=True)
    ap.add_argument("--src-dir", default=".")
    ap.add_argument("--docx-dir", default="_covers")
    ap.add_argument("--cover-dir", default="_covers")
    ap.add_argument("--out", default="fulltext_audited")
    args = ap.parse_args()

    data = json.load(open(args.meta, encoding="utf-8"))
    batch = data["batch"]
    batch.setdefault("total", len(data["items"]))
    items = data["items"]
    src_root = os.path.realpath(args.src_dir)
    for it in items:
        loc = it.setdefault("local", {})
        src = loc.get("file") or ""
        loc["abs"] = os.path.abspath(os.path.join(args.src_dir, src)) if src else ""
        # 路径护栏：清单（可能来自第三方）不能把读取口带到 src_dir 之外
        if loc["abs"]:
            real = os.path.realpath(loc["abs"])
            if not (real == src_root or real.startswith(src_root + os.sep)):
                raise SystemExit(f"[拒绝] 第 {it.get('no')} 条的 file 指向源目录之外：{src}")
            loc["abs"] = real
        # 页码区间必须是正序正整数
        ext = loc.get("extract")
        if ext:
            if (not isinstance(ext, (list, tuple)) or len(ext) != 2
                    or not all(isinstance(x, int) and x >= 1 for x in ext) or ext[0] > ext[1]):
                raise SystemExit(f"[拒绝] 第 {it.get('no')} 条的 extract 不是合法的 [起,止] 页码：{ext!r}")
        loc["exists"] = bool(src) and os.path.isfile(loc["abs"])

    if args.cmd == "build":
        from pypdf import PdfReader

        os.makedirs(args.docx_dir, exist_ok=True)
        made = 0
        for it in items:
            if not it["local"]["exists"]:
                continue
            loc = it["local"]
            p = loc["abs"]
            rdr = PdfReader(p)
            if loc.get("extract"):
                a, b = loc["extract"]
                loc["pages"] = b - a + 1
                loc["out_name"] = f"{it['no']:02d}.pdf"
            else:
                loc["pages"] = len(rdr.pages)
                loc["out_name"] = os.path.basename(p)
            loc["sha256"] = sha256_of(p)
            out = os.path.join(args.docx_dir, f"{it['no']:02d}.docx")
            build_docx(it, batch, out)
            made += 1
            print(f"[docx] {out}")
        print(f"covers built: {made}")
        return

    # merge
    from pypdf import PdfReader, PdfWriter

    os.makedirs(args.out, exist_ok=True)
    index = []
    for it in items:
        loc = it["local"]
        no = it["no"]
        if not loc["exists"]:
            index.append((no, it["title"], it["verdict"], "无正文", "", ""))
            continue
        cover = os.path.join(args.cover_dir, f"{no:02d}.pdf")
        if not os.path.isfile(cover):
            index.append((no, it["title"], it["verdict"], f"缺少审核页 {cover}", "", ""))
            continue
        src = PdfReader(loc["abs"])
        if loc.get("extract"):
            a, b = loc["extract"]
            pages = src.pages[a - 1 : b]
        else:
            pages = src.pages
        w = PdfWriter()
        w.add_page(PdfReader(cover).pages[0])
        for pg in pages:
            w.add_page(pg)
        name = f"{no:02d}_{slug(it['title'])}_{'审核版' if True else ''}.pdf"
        dst = os.path.join(args.out, name)
        with open(dst, "wb") as f:
            w.write(f)
        body_pages = len(pages)
        sha = loc.get("sha256") or sha256_of(loc["abs"])
        index.append((no, it["title"], it["verdict"], f"1+{body_pages} 页", name, sha[:16]))
        print(f"[merge] {dst}  (审核页 1 + 正文 {body_pages})")

    tsv = os.path.join(args.out, "audit_index.tsv")
    with open(tsv, "w", encoding="utf-8") as f:
        f.write("序号\t题名\t核验\t文件页数\t文件\t源文件SHA256(前16)\n")
        for row in index:
            f.write("\t".join(str(x) for x in row) + "\n")
    print(f"[index] {tsv}")


if __name__ == "__main__":
    sys.exit(main())
