#!/usr/bin/env python3
"""
Generator & Kompilator Makalah Akademik Standar (penyusun-makalah)
Mendukung kompilasi DOCX (python-docx) & PDF (reportlab two-pass TOC),
Logo Default Simetris, Font Unicode Times New Roman, Penomoran Halaman Dinamis
(Romawi & Arab Kanan Atas / Bawah Tengah Awal Bab), dan Academic Quality Linter.
"""

cfg = {
    "page": {
        "paper_size": "A4",
        "width_cm": 21.0,
        "height_cm": 29.7,
        "margin_preset": "4-3-3-3",
        "margin_left_cm": 4.0,
        "margin_top_cm": 3.0,
        "margin_right_cm": 3.0,
        "margin_bottom_cm": 3.0,
    },
    "typography": {
        "font_family": "Times New Roman",
        "body_size_pt": 12,
        "chapter_title_size_pt": 12,
        "subheading_size_pt": 12,
        "cover_title_size_pt": 14,
        "table_font_size_pt": 11,
        "caption_source_size_pt": 10,
        "line_spacing": 1.5,
        "single_spacing": 1.0,
        "first_line_indent_cm": 1.25,
        "hanging_indent_cm": 1.25,
        "blockquote_indent_cm": 1.25,
    },
    "structure": {
        "max_heading_depth": 3,
        "min_sentences_per_paragraph": 3,
        "max_sentences_per_paragraph": 7,
        "min_references": 11,
        "min_discussion_words": 1400,
        "min_paragraphs_per_main_sub": 3,
        "cover_logo_size_cm": 5.0,
        "numbering_style": "decimal",
    },
    "table": {
        "style": "word_grid",
        "header_bg": "F2F2F2",
        "border_sz": 6,
        "cell_pad_v_dxa": 90,
        "cell_pad_h_dxa": 130,
    },
    "pagination": {
        "preliminary_format": "lowerRoman",
        "body_format": "decimal",
        "preliminary_pos": "bottom_center",
        "body_first_page_pos": "bottom_center",
        "body_other_page_pos": "bottom_center",
    },
}

import argparse
import html
import io
import json
import math
import os
import re
import shutil
import sys
import tempfile
import zipfile

from PIL import Image, ImageDraw
import docx
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Flowable,
    Image as RLImage,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.doctemplate import BaseDocTemplate, PageTemplate
from reportlab.platypus.frames import Frame


ID_LOWERCASE_WORDS = {
    "di", "ke", "dari", "dan", "atau", "serta", "pada", "dalam", "untuk",
    "dengan", "oleh", "terhadap", "secara", "sebagai", "yang", "bagi",
    "tentang", "antara", "melalui", "menuju", "tanpa", "hingga", "sampai",
}

ACADEMIC_ABBREV_PATTERN = re.compile(
    r"\b(Prof|Dr|Ir|H|Hj|M\.T|M\.Sc|M\.Kom|M\.Pd|M\.Si|M\.Hum|S\.T|S\.Kom|S\.Pd|S\.E|S\.H|"
    r"dkk|dll|dsb|hlm|cet|ed|vol|no|UU|No|Rp|Yth|a\.n|u\.p|et al|e\.g|i\.e)\.",
    re.IGNORECASE,
)


def apply_margin_preset(preset: str) -> None:
    raw_parts = [x.strip() for x in preset.replace("-", ",").split(",") if x.strip()]
    try:
        parts = [float(x) for x in raw_parts]
    except ValueError as exc:
        raise ValueError(f"Format margin tidak valid ('{preset}'). Gunakan 4 angka, contoh: 4-3-3-3.") from exc
    if len(parts) != 4 or any(p <= 0 or p > 10 for p in parts):
        raise ValueError(f"Margin harus berisi 4 angka positif (Kiri,Atas,Kanan,Bawah), diterima: '{preset}'.")
    cfg["page"]["margin_left_cm"] = parts[0]
    cfg["page"]["margin_top_cm"] = parts[1]
    cfg["page"]["margin_right_cm"] = parts[2]
    cfg["page"]["margin_bottom_cm"] = parts[3]
    cfg["page"]["margin_preset"] = f"{parts[0]:g}-{parts[1]:g}-{parts[2]:g}-{parts[3]:g}"


def to_roman_lower(n: int) -> str:
    if n <= 0:
        return "i"
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman_num = ""
    i = 0
    while n > 0:
        for _ in range(n // val[i]):
            roman_num += syb[i]
            n -= val[i]
        i += 1
    return roman_num


def get_sorted_references(blueprint: dict) -> list[str]:
    return sorted(
        blueprint.get("references", []),
        key=lambda x: re.sub(r"^[^A-Za-z0-9]+", "", x).lower(),
    )


def get_cover_metadata(blueprint: dict) -> dict:
    title = blueprint.get("title", "JUDUL MAKALAH").strip().upper()
    subtitle = (blueprint.get("subtitle") or "").strip()
    course = (blueprint.get("course") or "").strip()
    lecturer = (blueprint.get("lecturer") or "").strip()
    authors = blueprint.get("authors", [{"name": "Nama Mahasiswa", "id": "NIM"}])
    acad_year = str(blueprint.get("academic_year") or blueprint.get("year", "2026")).strip()
    if re.match(r"^[0-9]{4}/[0-9]{4}$", acad_year):
        acad_year_label = f"TAHUN AKADEMIK {acad_year}"
    else:
        acad_year_label = acad_year.upper()

    inst_name = (blueprint.get("institution") or "UNIVERSITAS").strip()
    city_name = (blueprint.get("city") or "").strip()
    show_city = city_name if (city_name and not inst_name.lower().endswith(city_name.lower())) else ""

    inst_lines = [
        x.strip().upper()
        for x in [
            blueprint.get("department", ""),
            blueprint.get("faculty", ""),
            inst_name,
            show_city,
            acad_year_label,
        ]
        if x and str(x).strip()
    ]
    return {
        "title": title,
        "subtitle": subtitle,
        "course": course,
        "lecturer": lecturer,
        "authors": authors,
        "inst_lines": inst_lines,
    }


def ensure_default_logo(target_path: str) -> str:
    """Buat aset logo universitas default simetris beresolusi tinggi (600x600 px) jika belum ada."""
    os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)
    if os.path.exists(target_path) and os.path.getsize(target_path) > 1024:
        return target_path

    size = 600
    cx, cy = size // 2, size // 2
    img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    navy = (20, 42, 78, 255)
    gold = (197, 155, 52, 255)
    light_gold = (244, 232, 193, 255)
    white = (255, 255, 255, 255)

    outer_r = 270
    inner_r = 238
    pts_outer = []
    pts_inner = []
    for i in range(100):
        angle = (2 * math.pi * i) / 100 - math.pi / 2
        wave = 14 * math.cos(5 * (angle + math.pi / 2))
        r_out = outer_r + wave
        r_in = inner_r + wave * 0.85
        pts_outer.append((cx + r_out * math.cos(angle), cy + r_out * math.sin(angle)))
        pts_inner.append((cx + r_in * math.cos(angle), cy + r_in * math.sin(angle)))

    draw.polygon(pts_outer, fill=navy, outline=gold)
    for w in range(6):
        pts_ring = [
            (
                cx + (outer_r - w + 14 * math.cos(5 * ((2 * math.pi * i) / 100))) * math.cos((2 * math.pi * i) / 100 - math.pi / 2),
                cy + (outer_r - w + 14 * math.cos(5 * ((2 * math.pi * i) / 100))) * math.sin((2 * math.pi * i) / 100 - math.pi / 2),
            )
            for i in range(100)
        ]
        draw.polygon(pts_ring, outline=gold)

    draw.polygon(pts_inner, fill=navy, outline=gold)
    draw.ellipse((cx - 190, cy - 190, cx + 190, cy + 190), fill=light_gold, outline=gold, width=8)
    draw.ellipse((cx - 172, cy - 172, cx + 172, cy + 172), fill=navy, outline=gold, width=4)

    star_cx, star_cy = cx, cy - 45
    star_pts = []
    for i in range(16):
        ang = i * math.pi / 8 - math.pi / 2
        r = 58 if i % 2 == 0 else 24
        star_pts.append((star_cx + r * math.cos(ang), star_cy + r * math.sin(ang)))
    draw.polygon(star_pts, fill=gold, outline=white)

    book_y = cy + 62
    left_page = [(cx - 4, book_y - 28), (cx - 95, book_y - 42), (cx - 95, book_y + 28), (cx - 4, book_y + 44)]
    right_page = [(cx + 4, book_y - 28), (cx + 95, book_y - 42), (cx + 95, book_y + 28), (cx + 4, book_y + 44)]
    draw.polygon(left_page, fill=white, outline=gold)
    draw.polygon(right_page, fill=white, outline=gold)
    for offset in (-12, 4, 20):
        draw.line([(cx - 80, book_y + offset - 10), (cx - 16, book_y + offset)], fill=navy, width=3)
        draw.line([(cx + 80, book_y + offset - 10), (cx + 16, book_y + offset)], fill=navy, width=3)

    ribbon = [(cx - 115, cy + 125), (cx + 115, cy + 125), (cx + 100, cy + 150), (cx - 100, cy + 150)]
    draw.polygon(ribbon, fill=gold, outline=white)

    img.save(target_path, "PNG")
    return target_path


def _cap_first_alpha(word: str) -> str:
    m = re.match(r"^([^A-Za-z]*)([A-Za-z])(.*)$", word)
    if not m:
        return word
    return m.group(1) + m.group(2).upper() + m.group(3).lower()


def format_title_case_id(text: str) -> str:
    """Format judul sub-bab menjadi Title Case sesuai EYD V."""
    text = text.strip()
    match = re.match(r"^([0-9]+(?:\.[0-9]+)*\.?|[A-Z]\.|[a-z]\.|[0-9]+\)|[a-z]\))\s+(.*)$", text)
    prefix = ""
    rest = text
    if match:
        prefix = match.group(1) + " "
        rest = match.group(2)

    alpha_only = re.sub(r"[^A-Za-z]", "", rest)
    words = rest.split()
    all_caps_heading = len(words) >= 2 and len(alpha_only) >= 4 and alpha_only.isupper()

    formatted = []
    for idx, word in enumerate(words):
        clean = re.sub(r"[^A-Za-z0-9/-]", "", word)
        if not all_caps_heading and len(clean) >= 2 and clean.isupper():
            formatted.append(word)
        elif idx > 0 and idx < len(words) - 1 and clean.lower() in ID_LOWERCASE_WORDS:
            formatted.append(word.lower())
        elif "-" in word:
            formatted.append("-".join(_cap_first_alpha(w) for w in word.split("-")))
        elif "/" in word:
            formatted.append("/".join(_cap_first_alpha(w) for w in word.split("/")))
        else:
            formatted.append(_cap_first_alpha(word))
    return prefix + " ".join(formatted)


def clean_inline_artifacts(text: str) -> str:
    """Bersihkan sintaks markdown non-cetak seperti tautan [teks](url) dan backtick."""
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text


def tokenize_inline_markdown(
    text: str,
    default_bold: bool = False,
    default_italic: bool = False,
) -> list[tuple[str, bool, bool]]:
    """
    Tokenize inline Markdown (***bold-italic***, **bold**, *italic*, serta *italic* di dalam **bold**).
    Mengembalikan daftar (segment_text, is_bold, is_italic).
    """
    cleaned = clean_inline_artifacts(text)
    result: list[tuple[str, bool, bool]] = []
    tokens = re.split(r"(\*\*\*[^*]+?\*\*\*|\*\*.*?\*\*|\*[^*]+?\*)", cleaned)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("***") and tok.endswith("***") and len(tok) > 6:
            result.append((tok[3:-3], True, True))
        elif tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            inner = tok[2:-2]
            sub_tokens = re.split(r"(\*[^*]+?\*)", inner)
            for st in sub_tokens:
                if not st:
                    continue
                if st.startswith("*") and st.endswith("*") and len(st) > 2:
                    result.append((st[1:-1], True, True))
                else:
                    result.append((st, True, default_italic))
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            result.append((tok[1:-1], default_bold, True))
        else:
            result.append((tok, default_bold, default_italic))
    return result


def _strip_all_md_markers(text: str) -> str:
    return re.sub(r"\*+", "", clean_inline_artifacts(text)).strip()


def _next_non_empty_line(lines: list[str], start_idx: int) -> str:
    j = start_idx
    while j < len(lines):
        s = lines[j].strip()
        if s:
            return s
        j += 1
    return ""


def _is_table_caption_ahead(lines: list[str], idx: int) -> bool:
    stripped = lines[idx].strip()
    if not re.match(r"^(\*\*|\*)?Tabel\s+[0-9]+(\.[0-9]+)*", stripped, re.IGNORECASE):
        return False
    nxt = _next_non_empty_line(lines, idx + 1)
    return nxt.startswith("|") and nxt.endswith("|")


def parse_markdown_blocks(md_text: str) -> list[dict]:
    """Parse naskah Markdown menjadi blok terstruktur untuk DOCX, PDF, dan Linter."""
    lines = md_text.replace("\r\n", "\n").split("\n")
    blocks: list[dict] = []
    i = 0
    n = len(lines)
    pending_table_caption = None

    while i < n:
        raw = lines[i]
        stripped = raw.strip()

        if not stripped:
            i += 1
            continue

        if re.match(r"^[-*_]{3,}$", stripped):
            i += 1
            continue

        h_match = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if h_match:
            hashes = len(h_match.group(1))
            title_raw = h_match.group(2).strip()
            if hashes == 1 or re.match(r"^BAB\s+[IVX0-9]+", title_raw, re.IGNORECASE):
                i += 1
                continue
            blocks.append({
                "type": "heading",
                "raw_level": hashes,
                "level": min(hashes, cfg["structure"]["max_heading_depth"]),
                "text": format_title_case_id(title_raw),
                "raw_text": title_raw,
            })
            i += 1
            continue

        if _is_table_caption_ahead(lines, i):
            pending_table_caption = _strip_all_md_markers(stripped)
            i += 1
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows = []
            for tl in table_lines:
                if re.match(r"^\|[\s:|-]+\|$", tl):
                    continue
                raw_cells = re.split(r"(?<!\\)\|", tl.strip("|"))
                cells = [c.replace("\\|", "|").strip() for c in raw_cells]
                rows.append(cells)

            source_text = ""
            if i < n and re.match(r"^(\*|\*\*)?Sumber\s*:", lines[i].strip(), re.IGNORECASE):
                source_text = _strip_all_md_markers(lines[i].strip())
                i += 1

            blocks.append({
                "type": "table",
                "caption": pending_table_caption or "Tabel Data Analisis",
                "rows": rows,
                "source": source_text,
            })
            pending_table_caption = None
            continue

        img_match = re.match(r"^!\[(.*?)\]\((.*?)\)$", stripped)
        if img_match:
            alt_text = img_match.group(1).strip()
            img_path = img_match.group(2).strip()
            i += 1
            caption_text = f"Gambar Ilustrasi {alt_text}".strip()
            source_text = ""
            if i < n and re.match(r"^(\*\*|\*)?Gambar\s+[0-9]+(\.[0-9]+)*", lines[i].strip(), re.IGNORECASE):
                caption_text = _strip_all_md_markers(lines[i].strip())
                i += 1
            if i < n and re.match(r"^(\*|\*\*)?Sumber\s*:", lines[i].strip(), re.IGNORECASE):
                source_text = _strip_all_md_markers(lines[i].strip())
                i += 1
            blocks.append({
                "type": "figure",
                "path": img_path,
                "alt": alt_text,
                "caption": caption_text,
                "source": source_text,
            })
            continue

        if stripped.startswith(">"):
            q_lines = []
            while i < n and lines[i].strip().startswith(">"):
                q_lines.append(re.sub(r"^>\s?", "", lines[i].strip()))
                i += 1
            q_text = re.sub(r'^["“”]+|["“”]+$', "", " ".join(q_lines).strip()).strip()
            blocks.append({"type": "blockquote", "text": q_text})
            continue

        list_match = re.match(r"^([0-9]+\.|[a-z]\.|[-*•])\s+(.*)$", stripped)
        if list_match:
            expanded = raw.replace("\t", "    ")
            leading_spaces = len(expanded) - len(expanded.lstrip(" "))
            indent_level = min(2, leading_spaces // 2)
            marker = list_match.group(1)
            item_lines = [list_match.group(2).strip()]
            i += 1
            while i < n:
                nxt = lines[i].strip()
                if (
                    not nxt
                    or nxt.startswith("#")
                    or nxt.startswith("|")
                    or nxt.startswith(">")
                    or nxt.startswith("![")
                    or _is_table_caption_ahead(lines, i)
                    or re.match(r"^([0-9]+\.|[a-z]\.|[-*•])\s+", nxt)
                ):
                    break
                item_lines.append(nxt)
                i += 1
            blocks.append({
                "type": "list_item",
                "marker": marker,
                "indent_level": indent_level,
                "text": " ".join(item_lines),
            })
            continue

        p_lines = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt.startswith("#")
                or nxt.startswith("|")
                or nxt.startswith(">")
                or nxt.startswith("![")
                or _is_table_caption_ahead(lines, i)
                or re.match(r"^([0-9]+\.|[a-z]\.|[-*•])\s+", nxt)
            ):
                break
            p_lines.append(nxt)
            i += 1
        blocks.append({"type": "paragraph", "text": " ".join(p_lines)})

    return blocks


def split_sentences(text: str) -> list[str]:
    """Pisahkan paragraf menjadi kalimat dengan mengabaikan singkatan akademik & desimal."""
    cleaned = ACADEMIC_ABBREV_PATTERN.sub(r"\1", text)
    cleaned = re.sub(r"([0-9])\.([0-9])", r"\1_\2", cleaned)
    return [s.strip() for s in re.split(r"[.!?]+(?:\s+|$)", cleaned) if len(s.strip()) > 3]


def _extract_heading_depth(raw_text: str, numbering_style: str) -> int:
    s = raw_text.strip()
    dec = re.match(r"^([0-9]+(?:\.[0-9]+)+)", s)
    if dec:
        return len(dec.group(1).split("."))
    if numbering_style != "decimal":
        if re.match(r"^[0-9]+\)", s) or re.match(r"^[ivx]+\)", s, re.IGNORECASE):
            return 4
        if re.match(r"^[a-z]\.", s):
            return 3
        if re.match(r"^[0-9]+\.", s):
            return 2
        if re.match(r"^[A-Z]\.", s):
            return 1
    return 2


def audit_makalah(
    blueprint: dict,
    chapter_blocks: dict[str, list[dict]],
    base_dir: str,
    missing_chapters: list[str] | None = None,
) -> dict:
    """Academic Quality Linter dua arah sesuai 5 standar makalah akademik."""
    warnings: list[str] = list(missing_chapters or [])
    stats = {
        "total_paragraphs": 0,
        "total_tables": 0,
        "total_figures": 0,
        "total_references": len(blueprint.get("references", [])),
    }

    numbering_style = blueprint.get("numbering_style", cfg["structure"]["numbering_style"])
    all_text_blobs: list[str] = []
    rm_count = 0
    tujuan_count = 0
    kesimpulan_count = 0
    saran_blobs: list[str] = []
    current_sub = ""
    discussion_words = 0
    main_sub_para_counts: dict[str, int] = {}

    for ch_key, blocks in chapter_blocks.items():
        is_bab_ii = bool(re.match(r"^BAB\s+(?:II|2)(?:\s|$)", ch_key.strip(), re.IGNORECASE)) or "PEMBAHASAN" in ch_key.upper()
        active_main_sub = ""
        for idx, b in enumerate(blocks):
            btype = b["type"]
            if btype == "heading":
                current_sub = b["text"].lower()
                if is_bab_ii and b["level"] == 2:
                    active_main_sub = b["text"]
                    main_sub_para_counts.setdefault(active_main_sub, 0)
                depth = _extract_heading_depth(b["raw_text"], numbering_style)
                if b["raw_level"] > cfg["structure"]["max_heading_depth"] or depth > cfg["structure"]["max_heading_depth"]:
                    warnings.append(
                        f"[{ch_key}] Sub-bab melebihi 3 tingkat ({b['raw_text']}) -> dinormalisasi ke tingkat 3."
                    )
            elif btype == "paragraph":
                stats["total_paragraphs"] += 1
                txt = b["text"]
                all_text_blobs.append(txt)
                if is_bab_ii:
                    discussion_words += len(txt.split())
                    if active_main_sub:
                        main_sub_para_counts[active_main_sub] += 1
                if "saran" in current_sub:
                    saran_blobs.append(txt.lower())

                # Cek setiap kalimat di dalam paragraf untuk larangan awal angka
                raw_sentences = re.split(r"(?<=[.!?])\s+", txt)
                for sent in raw_sentences:
                    if re.match(r"^[0-9]+\s+[a-zA-Z]", sent.strip()):
                        warnings.append(
                            f"[{ch_key}] Kalimat diawali angka ('{sent.strip()[:35]}...'). Gunakan huruf sesuai EYD V."
                        )

                # Kecualikan paragraf pengantar singkat sebelum list/tabel/gambar
                nxt_type = blocks[idx + 1]["type"] if idx + 1 < len(blocks) else None
                is_lead_in = txt.rstrip().endswith(":") or nxt_type in ("list_item", "table", "figure")
                sc = len(split_sentences(txt))
                if not is_lead_in and sc < cfg["structure"]["min_sentences_per_paragraph"]:
                    warnings.append(
                        f"[{ch_key}] Paragraf terlalu pendek ({sc} kalimat, standar 3-7 kalimat): '{txt[:45]}...'"
                    )
                elif sc > cfg["structure"]["max_sentences_per_paragraph"] + 2:
                    warnings.append(
                        f"[{ch_key}] Paragraf terlalu panjang ({sc} kalimat, standar 3-7 kalimat): '{txt[:45]}...'"
                    )
            elif btype == "list_item":
                all_text_blobs.append(b["text"])
                if b["indent_level"] == 0:
                    if "rumusan masalah" in current_sub:
                        rm_count += 1
                    elif "tujuan" in current_sub:
                        tujuan_count += 1
                    elif "kesimpulan" in current_sub:
                        kesimpulan_count += 1
                    elif "saran" in current_sub:
                        saran_blobs.append(b["text"].lower())
            elif btype == "blockquote":
                all_text_blobs.append(b["text"])
                if is_bab_ii:
                    discussion_words += len(b["text"].split())
            elif btype == "table":
                stats["total_tables"] += 1
                if not b.get("source"):
                    warnings.append(f"[{ch_key}] {b['caption']} belum mencantumkan 'Sumber: ...' di bawah tabel.")
            elif btype == "figure":
                stats["total_figures"] += 1
                img_rel = b.get("path", "")
                img_abs = img_rel if os.path.isabs(img_rel) else os.path.join(base_dir, img_rel)
                if not os.path.exists(img_abs):
                    warnings.append(f"[{ch_key}] File gambar tidak ditemukan untuk {b['caption']}: '{img_rel}'.")
                if not b.get("source"):
                    warnings.append(f"[{ch_key}] {b['caption']} belum mencantumkan 'Sumber: ...' di bawah gambar.")

    # Sinkronisasi Rumusan Masalah, Tujuan Penulisan, dan Kesimpulan
    bp_ps = blueprint.get("problem_statements", [])
    expected_rm = len(bp_ps) if bp_ps else rm_count
    if bp_ps and rm_count > 0 and len(bp_ps) != rm_count:
        warnings.append(
            f"[Koherensi] Jumlah `problem_statements` di blueprint.json ({len(bp_ps)}) berbeda dengan butir Rumusan Masalah ({rm_count})."
        )
    if expected_rm > 0 and tujuan_count > 0 and expected_rm != tujuan_count:
        warnings.append(
            f"[Koherensi] Jumlah Rumusan Masalah ({expected_rm}) tidak sama dengan Tujuan Penulisan ({tujuan_count})."
        )
    if expected_rm > 0 and kesimpulan_count > 0 and expected_rm != kesimpulan_count:
        warnings.append(
            f"[Koherensi] Jumlah Rumusan Masalah ({expected_rm}) tidak sama dengan butir Kesimpulan ({kesimpulan_count})."
        )

    # Cek 3 target pihak pada Saran (Bab III)
    if saran_blobs:
        saran_text = " ".join(saran_blobs)
        if not re.search(r"\b(instansi|regulator|pemerintah|kementerian|lembaga)\b", saran_text):
            warnings.append("[Saran] Belum mencantumkan sasaran eksplisit bagi Instansi/Regulator.")
        if not re.search(r"\b(masyarakat|praktisi|industri|pelaku|pengguna|perusahaan)\b", saran_text):
            warnings.append("[Saran] Belum mencantumkan sasaran eksplisit bagi Masyarakat/Praktisi.")
        if not re.search(r"\b(peneliti|penulis selanjutnya|penelitian mendatang|akademisi)\b", saran_text):
            warnings.append("[Saran] Belum mencantumkan sasaran eksplisit bagi Peneliti/Penulis Selanjutnya.")

    # Sinkronisasi Sitasi Dua Arah (References <-> Teks)
    full_body_text = " ".join(all_text_blobs)
    ref_surnames: list[str] = []
    for ref in blueprint.get("references", []):
        if re.search(r"wikipedia\.org|blogspot\.com|wordpress\.com", ref, re.IGNORECASE):
            warnings.append(f"[Referensi Terlarang] Ditemukan sumber non-ilmiah di Daftar Pustaka: '{ref[:50]}...'")

        clean_ref = _strip_all_md_markers(ref)
        author_part = clean_ref.split("(")[0].strip().rstrip(".")
        first_surname = author_part.split(",")[0].strip().rstrip(".")
        year_match = re.search(r"\(([12][0-9]{3}[a-z]?)\)", clean_ref)
        if first_surname:
            ref_surnames.append(first_surname.lower())
            # Cek word-boundary agar "Tan" tidak cocok dengan "kegiatan"
            found = bool(re.search(rf"\b{re.escape(first_surname)}\b", full_body_text, re.IGNORECASE))
            if not found and " " in first_surname:
                # Dukungan singkatan instansi (mis. Kementerian Perindustrian -> Kemenperin)
                if year_match and year_match.group(1) in full_body_text:
                    found = True
            if not found:
                yr_str = year_match.group(1) if year_match else "?"
                warnings.append(
                    f"[Sinkronisasi Sitasi] Entri Daftar Pustaka '{first_surname} ({yr_str})' belum ditemukan sitasinya di dalam teks."
                )

    # Arah Teks -> Daftar Pustaka (deteksi sitasi (Penulis, 2023) yang lupa masuk Daftar Pustaka)
    ref_blob = " ".join(_strip_all_md_markers(r).lower() for r in blueprint.get("references", []))
    inline_cites = re.findall(r"\(([A-Z][A-Za-zÀ-ÿ\s.&'-]{1,80}?),\s*([12][0-9]{3}[a-z]?)(?::\s*[0-9-]+)?\)", full_body_text)
    seen_missing_cites: set[tuple[str, str]] = set()
    for cite_author, cite_year in inline_cites:
        lead_author = re.split(r"\s+(?:&|dan|et\s+al\.?|dkk\.?)(?:\s+|$)|,", cite_author.strip(), flags=re.IGNORECASE)[0].strip().lower()
        key = (lead_author, cite_year)
        if len(lead_author) >= 2 and key not in seen_missing_cites and not re.search(rf"\b{re.escape(lead_author)}\b", ref_blob):
            seen_missing_cites.add(key)
            warnings.append(
                f"[Sinkronisasi Sitasi] Sitasi di teks '({cite_author.strip()}, {cite_year})' belum terdaftar di Daftar Pustaka."
            )

    # Cek kedalaman pembahasan Bab II
    min_disc_words = cfg["structure"].get("min_discussion_words", 1400)
    min_sub_paras = cfg["structure"].get("min_paragraphs_per_main_sub", 3)
    if discussion_words > 0 and discussion_words < min_disc_words:
        warnings.append(
            f"[Kedalaman Pembahasan] Total kata Bab II baru {discussion_words} kata (disarankan minimal {min_disc_words} kata agar pembahasan mendalam)."
        )
    for sub_lbl, p_cnt in main_sub_para_counts.items():
        if p_cnt < min_sub_paras:
            warnings.append(
                f"[Kedalaman Sub-bab] Sub-bab '{sub_lbl}' baru berisi {p_cnt} paragraf (minimal {min_sub_paras} paragraf analitis agar seimbang)."
            )

    # Validasi keterbukaan referensi & jumlah minimal (> 10 / minimal 11 entri)
    min_refs = cfg["structure"].get("min_references", 11)
    refs_list = blueprint.get("references", [])
    if len(refs_list) < min_refs:
        warnings.append(
            f"[Jumlah Referensi] Daftar Pustaka baru berisi {len(refs_list)} entri (wajib di atas 10 / minimal {min_refs} referensi)."
        )

    verified_path = os.path.join(base_dir, "references_verified.json")
    if os.path.exists(verified_path):
        try:
            with open(verified_path, "r", encoding="utf-8") as vf:
                v_items = json.load(vf)
            open_cnt = 0
            for vi in v_items:
                if vi.get("status") != "OPEN_ACCESS":
                    warnings.append(
                        f"[Akses Referensi] Referensi '{vi.get('citation') or vi.get('url')}' berstatus {vi.get('status')} (bukan OPEN_ACCESS)."
                    )
                else:
                    open_cnt += 1
            if open_cnt < min_refs:
                warnings.append(
                    f"[Jumlah Referensi] Sumber OPEN_ACCESS terverifikasi baru {open_cnt} (wajib di atas 10 / minimal {min_refs} referensi)."
                )
        except Exception as exc:
            warnings.append(f"[Akses Referensi] Gagal membaca references_verified.json: {exc}")

    # Validasi kelengkapan komponen sampul (cegah placeholder '[...]' lolos)
    for fld in ("title", "course", "lecturer", "department", "faculty", "institution", "city"):
        val = str(blueprint.get(fld, "") or "").strip()
        if re.search(r"\[[^\]]+\]", val):
            warnings.append(
                f"[Komponen Sampul] Field '{fld}' masih berisi placeholder '{val}'. Minta pengguna mengisi via chat."
            )
    for idx_a, auth in enumerate(blueprint.get("authors", []), 1):
        for af in ("name", "id"):
            aval = str(auth.get(af, "") or "").strip()
            if re.search(r"\[[^\]]+\]", aval) or (af == "name" and aval.lower() == "nama mahasiswa"):
                warnings.append(
                    f"[Komponen Sampul] Penulis #{idx_a} '{af}' masih berisi placeholder '{aval}'. Minta pengguna mengisi via chat."
                )

    return {"stats": stats, "warnings": warnings}


# ==============================================================================
# DOCX GENERATION ENGINE (python-docx)
# ==============================================================================

def add_xml_page_number(run) -> None:
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def set_section_pg_num_type(section, fmt: str = "decimal", start: int | None = None) -> None:
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        sectPr.append(pgNumType)
    pgNumType.set(qn("w:fmt"), fmt)
    if start is not None:
        pgNumType.set(qn("w:start"), str(start))
    elif qn("w:start") in pgNumType.attrib:
        del pgNumType.attrib[qn("w:start")]


def clear_story_paragraphs(story) -> None:
    for p in list(story.paragraphs):
        p.text = ""


def configure_section_margins(section) -> None:
    section.page_width = Cm(cfg["page"]["width_cm"])
    section.page_height = Cm(cfg["page"]["height_cm"])
    section.left_margin = Cm(cfg["page"]["margin_left_cm"])
    section.top_margin = Cm(cfg["page"]["margin_top_cm"])
    section.right_margin = Cm(cfg["page"]["margin_right_cm"])
    section.bottom_margin = Cm(cfg["page"]["margin_bottom_cm"])
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.5)


def _set_story_page_number(story_obj, alignment) -> None:
    clear_story_paragraphs(story_obj)
    p = story_obj.paragraphs[0]
    p.alignment = alignment
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    _apply_run_font(run, cfg["typography"]["body_size_pt"], False, False)
    add_xml_page_number(run)


def setup_preliminary_pagination(section, restart_at_one: bool = False) -> None:
    configure_section_margins(section)
    section.different_first_page_header_footer = False
    set_section_pg_num_type(section, fmt="lowerRoman", start=1 if restart_at_one else None)

    hdr = section.header
    hdr.is_linked_to_previous = False
    clear_story_paragraphs(hdr)

    ftr = section.footer
    ftr.is_linked_to_previous = False
    _set_story_page_number(ftr, WD_ALIGN_PARAGRAPH.CENTER)


def setup_body_chapter_pagination(section, restart_at_one: bool = False) -> None:
    configure_section_margins(section)
    set_section_pg_num_type(section, fmt="decimal", start=1 if restart_at_one else None)
    other_pos = cfg["pagination"].get("body_other_page_pos", "bottom_center")

    if other_pos == "bottom_center":
        section.different_first_page_header_footer = False
        hdr = section.header
        hdr.is_linked_to_previous = False
        clear_story_paragraphs(hdr)
        ftr = section.footer
        ftr.is_linked_to_previous = False
        _set_story_page_number(ftr, WD_ALIGN_PARAGRAPH.CENTER)
        return

    section.different_first_page_header_footer = True
    fp_hdr = section.first_page_header
    fp_hdr.is_linked_to_previous = False
    clear_story_paragraphs(fp_hdr)

    fp_ftr = section.first_page_footer
    fp_ftr.is_linked_to_previous = False
    _set_story_page_number(fp_ftr, WD_ALIGN_PARAGRAPH.CENTER)

    hdr = section.header
    hdr.is_linked_to_previous = False
    _set_story_page_number(hdr, WD_ALIGN_PARAGRAPH.RIGHT)

    ftr = section.footer
    ftr.is_linked_to_previous = False
    clear_story_paragraphs(ftr)


def _apply_run_font(run, font_size_pt: float, is_bold: bool, is_italic: bool) -> None:
    font_name = cfg["typography"]["font_family"]
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.bold = is_bold
    run.italic = is_italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font_name)


def add_formatted_runs(
    paragraph,
    text: str,
    font_size_pt: float = 12,
    default_bold: bool = False,
    default_italic: bool = False,
) -> None:
    for seg, is_bold, is_italic in tokenize_inline_markdown(text, default_bold, default_italic):
        run = paragraph.add_run(seg)
        _apply_run_font(run, font_size_pt, is_bold, is_italic)


def _set_cell_shading(cell, fill_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is not None:
        tcPr.remove(shd)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def _compute_col_widths_cm(rows: list[list[str]], total_width_cm: float) -> list[float]:
    num_cols = max(len(r) for r in rows)
    max_lens = [4] * num_cols
    for r in rows:
        for c_idx in range(num_cols):
            val = r[c_idx] if c_idx < len(r) else ""
            max_lens[c_idx] = max(max_lens[c_idx], min(48, len(val)))
    total_chars = sum(max_lens) or num_cols
    raw = [max(total_width_cm * 0.10, total_width_cm * (ln / total_chars)) for ln in max_lens]
    scale = total_width_cm / sum(raw)
    return [w * scale for w in raw]


def style_academic_table(table, col_widths_cm: list[float]) -> None:
    """Format tabel standar Microsoft Word (Table Grid penuh, lebar penuh margin, padding rapi, header berulang)."""
    tbl_cfg = cfg["table"]
    is_grid = tbl_cfg.get("style", "word_grid") == "word_grid"
    border_sz = str(tbl_cfg.get("border_sz", 6))

    table.autofit = False
    tblPr = table._tbl.tblPr

    # Lebar tabel tetap 100% area cetak + layout fixed
    total_dxa = int(round(sum(col_widths_cm) * 566.929))
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is not None:
        tblPr.remove(tblW)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is not None:
        tblPr.remove(tblLayout)
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    # Cell padding standar Word yang lega
    tblCellMar = tblPr.find(qn("w:tblCellMar"))
    if tblCellMar is not None:
        tblPr.remove(tblCellMar)
    tblCellMar = OxmlElement("w:tblCellMar")
    pad_v = str(tbl_cfg.get("cell_pad_v_dxa", 90))
    pad_h = str(tbl_cfg.get("cell_pad_h_dxa", 130))
    for side, val in [("top", pad_v), ("bottom", pad_v), ("left", pad_h), ("right", pad_h)]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val)
        m.set(qn("w:type"), "dxa")
        tblCellMar.append(m)
    tblPr.append(tblCellMar)

    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    tblBorders = OxmlElement("w:tblBorders")
    if is_grid:
        border_specs = [(side, "single", border_sz) for side in ("top", "bottom", "left", "right", "insideH", "insideV")]
    else:
        border_specs = [
            ("top", "single", "12"), ("bottom", "single", "12"),
            ("left", "none", "0"), ("right", "none", "0"),
            ("insideH", "none", "0"), ("insideV", "none", "0"),
        ]
    for border_name, val, sz in border_specs:
        b = OxmlElement(f"w:{border_name}")
        b.set(qn("w:val"), val)
        b.set(qn("w:sz"), sz)
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    tblPr.append(tblBorders)

    for c_idx, col_w in enumerate(col_widths_cm):
        table.columns[c_idx].width = Cm(col_w)

    header_bg = tbl_cfg.get("header_bg", "F2F2F2")
    for r_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if r_idx == 0:
            trPr.append(OxmlElement("w:tblHeader"))
        for c_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_cm[c_idx])
            if r_idx == 0:
                if is_grid and header_bg:
                    _set_cell_shading(cell, header_bg)
                elif not is_grid:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcBorders = OxmlElement("w:tcBorders")
                    bottom = OxmlElement("w:bottom")
                    bottom.set(qn("w:val"), "single")
                    bottom.set(qn("w:sz"), "8")
                    bottom.set(qn("w:space"), "0")
                    bottom.set(qn("w:color"), "000000")
                    tcBorders.append(bottom)
                    tcPr.append(tcBorders)


def resolve_cover_logo(blueprint: dict, base_dir: str, fallback_logo: str) -> str:
    """Cari file foto logo kampus dari blueprint.json atau folder assets/, fallback ke logo default."""
    user_logo = blueprint.get("logo_path", "").strip()
    if user_logo:
        cand = user_logo if os.path.isabs(user_logo) else os.path.join(base_dir, user_logo)
        if os.path.exists(cand):
            return cand
    assets_dir = os.path.join(base_dir, "assets")
    for name in (
        "logo_kampus.png", "logo_kampus.jpg", "logo_kampus.jpeg", "logo_kampus.webp",
        "logo.png", "logo.jpg", "logo.jpeg",
    ):
        cand = os.path.join(assets_dir, name)
        if os.path.exists(cand):
            return cand
    return fallback_logo


def compute_cover_dimensions_cm(logo_path: str, max_dim_cm: float) -> tuple[float, float]:
    """Pertahankan rasio asli foto logo kampus agar tidak terdistorsi."""
    if not logo_path or not os.path.exists(logo_path):
        return max_dim_cm, max_dim_cm
    with Image.open(logo_path) as im:
        w_px, h_px = im.size
    if w_px <= 0 or h_px <= 0:
        return max_dim_cm, max_dim_cm
    if w_px >= h_px:
        return max_dim_cm, max_dim_cm * (h_px / w_px)
    return max_dim_cm * (w_px / h_px), max_dim_cm


def _measure_flow_list_h(flows: list, avail_w: float, avail_h: float) -> float:
    """Hitung total tinggi ter-render (pt) dari daftar Flowable ReportLab termasuk spaceAfter antar-item."""
    total = 0.0
    for i, fl in enumerate(flows):
        _, fh = fl.wrap(avail_w, avail_h)
        total += fh
        if i < len(flows) - 1 and hasattr(fl, "style"):
            sa = getattr(fl.style, "spaceAfter", 0) or 0
            nxt_sb = getattr(getattr(flows[i + 1], "style", None), "spaceBefore", 0) or 0
            total += max(sa, nxt_sb)
    return total


def _build_cover_zone_spec(blueprint: dict, logo_path: str, fonts: tuple[str, str, str, str] | None = None) -> dict:
    """Bangun 3 zona sampul (Atas, Tengah, Bawah Rapat Margin) beserta tinggi presisi untuk paritas DOCX dan PDF."""
    f_reg, f_bold, _, _ = fonts or register_pdf_fonts()
    styles = getSampleStyleSheet()
    meta = get_cover_metadata(blueprint)
    authors = meta["authors"]
    has_logo = bool(logo_path and os.path.exists(logo_path))

    cov_frame_w = (cfg["page"]["width_cm"] - 6.0) * cm
    usable_h = (cfg["page"]["height_cm"] - cfg["page"]["margin_top_cm"] - cfg["page"]["margin_bottom_cm"]) * cm
    cov_frame_h = usable_h - 6.0

    st_lbl = ParagraphStyle("CovLbl", parent=styles["Normal"], fontName=f_bold, fontSize=16, leading=20, alignment=TA_CENTER, spaceAfter=14)
    st_ttl = ParagraphStyle("CovTtl", parent=styles["Normal"], fontName=f_bold, fontSize=14, leading=19, alignment=TA_CENTER, spaceAfter=10)
    st_sub = ParagraphStyle("CovSub", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=4)
    st_meta = ParagraphStyle("CovMeta", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=3)
    st_meta_last = ParagraphStyle("CovMetaLast", parent=st_meta, spaceAfter=0)
    st_by = ParagraphStyle("CovBy", parent=styles["Normal"], fontName=f_bold, fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=5)
    st_auth = ParagraphStyle("CovAuth", parent=styles["Normal"], fontName=f_bold, fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=2)
    st_auth_last = ParagraphStyle("CovAuthLast", parent=st_auth, spaceAfter=0)
    st_inst = ParagraphStyle("CovInst", parent=styles["Normal"], fontName=f_bold, fontSize=13, leading=17.5, alignment=TA_CENTER, spaceAfter=2)
    st_inst_last = ParagraphStyle("CovInstLast", parent=st_inst, spaceAfter=0)

    top_flows: list = [Paragraph("MAKALAH", st_lbl)]
    has_meta_below_title = bool(meta["subtitle"] or meta["course"] or meta["lecturer"])
    top_flows.append(
        Paragraph(
            md_inline_to_rl_xml(meta["title"], default_bold=True),
            st_ttl if has_meta_below_title else ParagraphStyle("CovTtl0", parent=st_ttl, spaceAfter=0),
        )
    )
    if meta["subtitle"]:
        top_flows.append(Paragraph(md_inline_to_rl_xml(meta["subtitle"]), st_sub if meta["lecturer"] else st_meta_last))
    elif meta["course"]:
        top_flows.append(Paragraph(md_inline_to_rl_xml(f"Mata Kuliah: {meta['course']}"), st_meta if meta["lecturer"] else st_meta_last))
    if meta["lecturer"]:
        top_flows.append(Paragraph(md_inline_to_rl_xml(f"Dosen Pengampu: {meta['lecturer']}"), st_meta_last))

    bot_flows: list = []
    for idx_i, iline in enumerate(meta["inst_lines"]):
        st_use = st_inst_last if idx_i == len(meta["inst_lines"]) - 1 else st_inst
        bot_flows.append(Paragraph(html.escape(iline), st_use))

    def _make_mid(logo_cap_cm: float) -> tuple[list, float, float]:
        lw, lh = compute_cover_dimensions_cm(logo_path, logo_cap_cm)
        m_flows: list = []
        if has_logo and lw > 0 and lh > 0:
            im = RLImage(logo_path, width=lw * cm, height=lh * cm, mask="auto")
            im.hAlign = "CENTER"
            m_flows.append(im)
            m_flows.append(Spacer(1, 20))
        m_flows.append(Paragraph("Disusun Oleh:", st_by))
        if len(authors) > 1:
            auth_rows = []
            st_al = ParagraphStyle("CovAL", parent=st_auth_last, alignment=TA_LEFT, leading=15)
            st_ar = ParagraphStyle("CovAR", parent=st_auth_last, alignment=TA_RIGHT, leading=15)
            for a in authors:
                nm = html.escape(a.get("name", ""))
                nim = f"NIM. {html.escape((a.get('id') or '').strip())}" if (a.get("id") or "").strip() else ""
                auth_rows.append([Paragraph(nm, st_al), Paragraph(nim, st_ar)])
            t_auth = Table(auth_rows, colWidths=[6.5 * cm, 4.0 * cm])
            t_auth.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
            t_auth.hAlign = "CENTER"
            m_flows.append(t_auth)
        else:
            a = authors[0] if authors else {"name": "", "id": ""}
            nm = html.escape(a.get("name", ""))
            nim = html.escape((a.get("id") or "").strip())
            m_flows.append(Paragraph(nm, st_auth if nim else st_auth_last))
            if nim:
                m_flows.append(Paragraph(f"NIM. {nim}", st_auth_last))
        return m_flows, lw, lh

    h_top = _measure_flow_list_h(top_flows, cov_frame_w, cov_frame_h)
    h_bot = _measure_flow_list_h(bot_flows, cov_frame_w, cov_frame_h)

    base_cap = 4.0 if len(authors) > 3 else cfg["structure"]["cover_logo_size_cm"]
    mid_flows, logo_w_cm, logo_h_cm = _make_mid(base_cap)
    h_mid = _measure_flow_list_h(mid_flows, cov_frame_w, cov_frame_h)

    for fallback_cap in (4.2, 3.6, 3.0):
        if h_top + h_mid + h_bot + 28.0 <= cov_frame_h or not has_logo:
            break
        if fallback_cap < base_cap:
            mid_flows, logo_w_cm, logo_h_cm = _make_mid(fallback_cap)
            h_mid = _measure_flow_list_h(mid_flows, cov_frame_w, cov_frame_h)

    spare = max(18.0, cov_frame_h - (h_top + h_mid + h_bot))
    r1_h = h_top + spare * 0.25
    r2_h = h_mid + spare * 0.50
    r3_h = max(h_bot + 4.0, cov_frame_h - r1_h - r2_h)
    return {
        "meta": meta,
        "has_logo": has_logo,
        "logo_w_cm": logo_w_cm,
        "logo_h_cm": logo_h_cm,
        "cov_frame_w": cov_frame_w,
        "cov_frame_h": cov_frame_h,
        "top_flows": top_flows,
        "mid_flows": mid_flows,
        "bot_flows": bot_flows,
        "row_heights": (r1_h, r2_h, r3_h),
    }


def set_borderless_table(table) -> None:
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ("top", "bottom", "left", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{border_name}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def _set_style_rfonts(style, font_name: str) -> None:
    rPr = style._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font_name)


def configure_docx_styles(doc) -> None:
    """Konfigurasikan Normal, Heading 1-3, Caption Tabel/Gambar, dan TOC 1-3 agar selaras dengan Word TOC."""
    font_name = cfg["typography"]["font_family"]
    printable_width_cm = (
        cfg["page"]["width_cm"]
        - cfg["page"]["margin_left_cm"]
        - cfg["page"]["margin_right_cm"]
    )

    normal_style = doc.styles["Normal"]
    normal_style.font.name = font_name
    normal_style.font.size = Pt(cfg["typography"]["body_size_pt"])
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.line_spacing = cfg["typography"]["line_spacing"]
    _set_style_rfonts(normal_style, font_name)

    heading_specs = [
        ("Heading 1", WD_ALIGN_PARAGRAPH.CENTER, 0, 18, 0),
        ("Heading 2", WD_ALIGN_PARAGRAPH.LEFT, 12, 4, 0),
        ("Heading 3", WD_ALIGN_PARAGRAPH.LEFT, 8, 4, 0),
    ]
    for st_name, align, sb, sa, indent in heading_specs:
        st = doc.styles[st_name]
        st.font.name = font_name
        st.font.size = Pt(12)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        pf = st.paragraph_format
        pf.alignment = align
        pf.left_indent = Cm(indent)
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(sb)
        pf.space_after = Pt(sa)
        pf.line_spacing = cfg["typography"]["line_spacing"]
        pf.keep_with_next = True
        pf.keep_together = True
        _set_style_rfonts(st, font_name)

    for cap_name, align in [("Caption Tabel", WD_ALIGN_PARAGRAPH.LEFT), ("Caption Gambar", WD_ALIGN_PARAGRAPH.CENTER)]:
        if cap_name not in doc.styles:
            cap_st = doc.styles.add_style(cap_name, WD_STYLE_TYPE.PARAGRAPH)
        else:
            cap_st = doc.styles[cap_name]
        cap_st.base_style = doc.styles["Normal"]
        cap_st.font.name = font_name
        cap_st.font.size = Pt(12)
        cap_st.font.bold = True
        cap_st.font.color.rgb = RGBColor(0, 0, 0)
        pf = cap_st.paragraph_format
        pf.alignment = align
        pf.first_line_indent = Cm(0)
        pf.line_spacing = 1.0
        pf.keep_with_next = True
        _set_style_rfonts(cap_st, font_name)

    for toc_idx, (toc_name, indent_cm, is_bold) in enumerate([
        ("TOC 1", 0.0, True),
        ("TOC 2", 0.75, False),
        ("TOC 3", 1.5, False),
    ], 1):
        if toc_name not in doc.styles:
            toc_st = doc.styles.add_style(toc_name, WD_STYLE_TYPE.PARAGRAPH)
        else:
            toc_st = doc.styles[toc_name]
        toc_st.base_style = doc.styles["Normal"]
        toc_st.font.name = font_name
        toc_st.font.size = Pt(12)
        toc_st.font.bold = is_bold
        toc_st.font.color.rgb = RGBColor(0, 0, 0)
        pf = toc_st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.left_indent = Cm(indent_cm)
        pf.right_indent = Cm(1.0)
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        pf.line_spacing = cfg["typography"]["line_spacing"]
        pf.tab_stops.clear_all()
        pf.tab_stops.add_tab_stop(Cm(printable_width_cm), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        _set_style_rfonts(toc_st, font_name)


def enable_update_fields(doc) -> None:
    """Aktifkan updateFields=true di word/settings.xml agar Word memperbarui TOC otomatis saat dibuka."""
    settings_el = doc.settings.element
    uf = settings_el.find(qn("w:updateFields"))
    if uf is None:
        uf = OxmlElement("w:updateFields")
        settings_el.append(uf)
    uf.set(qn("w:val"), "true")


def _build_toc_entry_paragraph(doc, label: str, page_str: str, level: int, bold: bool) -> OxmlElement:
    style_name = f"TOC {min(3, max(1, level))}"
    p = doc.add_paragraph(style=style_name)
    add_formatted_runs(p, label, font_size_pt=12, default_bold=bold)
    r_tab = p.add_run(f"\t{page_str}")
    _apply_run_font(r_tab, 12, bold, False)
    return p._p


def add_native_word_toc(
    doc,
    field_instr: str,
    cached_entries: list[dict],
    gallery_name: str = "Table of Contents",
) -> None:
    """
    Sisipkan blok Table of Contents native Microsoft Word (<w:sdt> + field begin/separate/end)
    dengan entri cached dari two-pass rendering sehingga langsung tampil rapi & bisa di-Update Table.
    """
    body = doc._body._body
    sdt = OxmlElement("w:sdt")

    sdtPr = OxmlElement("w:sdtPr")
    docPartObj = OxmlElement("w:docPartObj")
    docPartGallery = OxmlElement("w:docPartGallery")
    docPartGallery.set(qn("w:val"), gallery_name)
    docPartUnique = OxmlElement("w:docPartUnique")
    docPartObj.append(docPartGallery)
    docPartObj.append(docPartUnique)
    sdtPr.append(docPartObj)
    sdt.append(sdtPr)

    sdtContent = OxmlElement("w:sdtContent")

    # Paragraf pembuka field TOC (begin -> instrText -> separate)
    p_start = doc.add_paragraph()
    p_start.paragraph_format.first_line_indent = Cm(0)
    p_start.paragraph_format.space_before = Pt(0)
    p_start.paragraph_format.space_after = Pt(0)
    r_begin = p_start.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    fld_begin.set(qn("w:dirty"), "true")
    r_begin._r.append(fld_begin)

    r_instr = p_start.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_instr} "
    r_instr._r.append(instr)

    r_sep = p_start.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep._r.append(fld_sep)

    # Pindahkan p_start ke dalam sdtContent
    body.remove(p_start._p)
    sdtContent.append(p_start._p)

    # Tambahkan entri TOC hasil two-pass di antara separate & end
    for item in cached_entries:
        indent_cm = item.get("indent_cm", 0.0)
        lvl = 1 if indent_cm == 0.0 else (2 if indent_cm <= 0.8 else 3)
        p_el = _build_toc_entry_paragraph(doc, item["label"], item["page"], lvl, item.get("bold", lvl == 1))
        body.remove(p_el)
        sdtContent.append(p_el)

    # Paragraf penutup field TOC (end)
    p_end = doc.add_paragraph()
    p_end.paragraph_format.first_line_indent = Cm(0)
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    r_end = p_end.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end._r.append(fld_end)
    body.remove(p_end._p)
    sdtContent.append(p_end._p)

    sdt.append(sdtContent)
    body.append(sdt)


def add_toc_line(doc, label: str, page_str: str, indent_cm: float = 0.0, bold: bool = False) -> None:
    lvl = 1 if indent_cm == 0.0 else (2 if indent_cm <= 0.8 else 3)
    _build_toc_entry_paragraph(doc, label, page_str, lvl, bold)


def _render_docx_blocks(doc, blocks: list[dict], base_dir: str) -> None:
    for b in blocks:
        btype = b["type"]
        if btype == "heading":
            h_style = "Heading 2" if b["level"] <= 2 else "Heading 3"
            p_h = doc.add_paragraph(style=h_style)
            add_formatted_runs(p_h, b["text"], font_size_pt=12, default_bold=True)
        elif btype == "paragraph":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf = p.paragraph_format
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(cfg["typography"]["first_line_indent_cm"])
            pf.space_before = Pt(0)
            pf.space_after = Pt(6)
            pf.line_spacing = cfg["typography"]["line_spacing"]
            pf.widow_control = True
            add_formatted_runs(p, b["text"], font_size_pt=12)
        elif btype == "blockquote":
            p_q = doc.add_paragraph()
            p_q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf = p_q.paragraph_format
            pf.left_indent = Cm(cfg["typography"]["blockquote_indent_cm"])
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(6)
            pf.line_spacing = cfg["typography"]["single_spacing"]
            pf.widow_control = True
            add_formatted_runs(p_q, b["text"], font_size_pt=12)
        elif btype == "list_item":
            p_li = doc.add_paragraph()
            p_li.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf = p_li.paragraph_format
            extra_indent = 0.5 * b.get("indent_level", 0)
            pf.left_indent = Cm(0.75 + extra_indent)
            pf.first_line_indent = Cm(-0.75)
            pf.space_before = Pt(0)
            pf.space_after = Pt(3)
            pf.line_spacing = cfg["typography"]["line_spacing"]
            pf.widow_control = True
            add_formatted_runs(p_li, f"{b['marker']} {b['text']}", font_size_pt=12)
        elif btype == "table":
            p_tcap = doc.add_paragraph(style="Caption Tabel")
            pf_cap = p_tcap.paragraph_format
            pf_cap.space_before = Pt(10)
            pf_cap.space_after = Pt(4)
            add_formatted_runs(p_tcap, b["caption"], font_size_pt=12, default_bold=True)

            rows = b.get("rows", [])
            if rows:
                num_cols = max(len(r) for r in rows)
                norm_rows = [r + [""] * (num_cols - len(r)) for r in rows]
                printable_w_cm = (
                    cfg["page"]["width_cm"]
                    - cfg["page"]["margin_left_cm"]
                    - cfg["page"]["margin_right_cm"]
                )
                col_widths_cm = _compute_col_widths_cm(norm_rows, printable_w_cm)
                tbl = doc.add_table(rows=len(norm_rows), cols=num_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                for r_idx, r_data in enumerate(norm_rows):
                    for c_idx in range(num_cols):
                        val = r_data[c_idx]
                        cell = tbl.cell(r_idx, c_idx)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        cp = cell.paragraphs[0]
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER if r_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                        cp.paragraph_format.first_line_indent = Cm(0)
                        cp.paragraph_format.space_before = Pt(2)
                        cp.paragraph_format.space_after = Pt(2)
                        cp.paragraph_format.line_spacing = 1.15
                        add_formatted_runs(
                            cp,
                            val,
                            font_size_pt=cfg["typography"]["table_font_size_pt"],
                            default_bold=(r_idx == 0),
                        )
                style_academic_table(tbl, col_widths_cm)

            if b.get("source"):
                p_tsrc = doc.add_paragraph()
                p_tsrc.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf_src = p_tsrc.paragraph_format
                pf_src.first_line_indent = Cm(0)
                pf_src.space_before = Pt(4)
                pf_src.space_after = Pt(10)
                pf_src.line_spacing = 1.0
                add_formatted_runs(p_tsrc, b["source"], font_size_pt=10, default_italic=True)
        elif btype == "figure":
            img_rel = b.get("path", "")
            img_abs = img_rel if os.path.isabs(img_rel) else os.path.join(base_dir, img_rel)
            if os.path.exists(img_abs):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.first_line_indent = Cm(0)
                p_img.paragraph_format.space_before = Pt(10)
                p_img.paragraph_format.space_after = Pt(4)
                p_img.paragraph_format.keep_with_next = True
                r_img = p_img.add_run()
                with Image.open(img_abs) as im_check:
                    w_px, h_px = im_check.size
                max_w_cm = 11.0
                max_h_cm = 12.0
                ratio = h_px / max(1, w_px)
                target_w = max_w_cm
                if target_w * ratio > max_h_cm:
                    target_w = max_h_cm / ratio
                r_img.add_picture(img_abs, width=Cm(target_w))

            p_fcap = doc.add_paragraph(style="Caption Gambar")
            pf_fcap = p_fcap.paragraph_format
            pf_fcap.space_before = Pt(2)
            pf_fcap.space_after = Pt(2)
            pf_fcap.keep_with_next = bool(b.get("source"))
            add_formatted_runs(p_fcap, b["caption"], font_size_pt=12, default_bold=True)

            if b.get("source"):
                p_fsrc = doc.add_paragraph()
                p_fsrc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_fsrc.paragraph_format.first_line_indent = Cm(0)
                p_fsrc.paragraph_format.space_before = Pt(0)
                p_fsrc.paragraph_format.space_after = Pt(10)
                p_fsrc.paragraph_format.line_spacing = 1.0
                add_formatted_runs(p_fsrc, b["source"], font_size_pt=10, default_italic=True)


def build_docx(
    blueprint: dict,
    abstract_blocks: list[dict],
    preface_blocks: list[dict],
    chapter_blocks: list[tuple[dict, list[dict]]],
    logo_path: str,
    base_dir: str,
    output_docx_path: str,
    toc_data: dict | None = None,
) -> None:
    doc = docx.Document()
    configure_docx_styles(doc)
    enable_update_fields(doc)

    # SECTION 0: HALAMAN SAMPUL (COVER) — Margin kiri-kanan simetris 3,0 cm agar logo & teks tepat di tengah
    sec_cover = doc.sections[0]
    configure_section_margins(sec_cover)
    sec_cover.left_margin = Cm(3.0)
    sec_cover.right_margin = Cm(3.0)
    sec_cover.different_first_page_header_footer = True
    clear_story_paragraphs(sec_cover.first_page_header)
    clear_story_paragraphs(sec_cover.first_page_footer)
    clear_story_paragraphs(sec_cover.header)
    clear_story_paragraphs(sec_cover.footer)

    spec = _build_cover_zone_spec(blueprint, logo_path)
    meta = spec["meta"]
    authors = meta["authors"]
    has_logo = spec["has_logo"]
    logo_w_cm, logo_h_cm = spec["logo_w_cm"], spec["logo_h_cm"]
    r1_h, r2_h, r3_h = spec["row_heights"]

    # Grid 3-Zona Penuh (Atas-Tengah-Bawah Rapat Margin) setara presisi dengan PDF
    tbl_cov = doc.add_table(rows=3, cols=1)
    tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_cov.autofit = False
    set_borderless_table(tbl_cov)
    tblPr_cov = tbl_cov._tbl.tblPr
    tblMar = tblPr_cov.find(qn("w:tblCellMar"))
    if tblMar is not None:
        tblPr_cov.remove(tblMar)
    tblMar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), "0")
        m.set(qn("w:type"), "dxa")
        tblMar.append(m)
    tblPr_cov.append(tblMar)

    cov_w_cm = Cm(15.0)
    row_aligns = (WD_ALIGN_VERTICAL.TOP, WD_ALIGN_VERTICAL.CENTER, WD_ALIGN_VERTICAL.BOTTOM)
    row_heights_pt = (max(20.0, r1_h - 3.0), max(20.0, r2_h - 3.0), max(20.0, r3_h - 3.0))
    for idx_r, (rh_pt, v_al) in enumerate(zip(row_heights_pt, row_aligns)):
        row = tbl_cov.rows[idx_r]
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Pt(rh_pt)
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        c = row.cells[0]
        c.width = cov_w_cm
        c.vertical_alignment = v_al

    def _add_cov_p(cell, idx_p: int, text: str, pt_size: int, bold: bool = False, sa: float = 2.0, ls: float = 1.25):
        p = cell.paragraphs[0] if idx_p == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(sa)
        pf.line_spacing = ls
        add_formatted_runs(p, text, font_size_pt=pt_size, default_bold=bold)
        return p

    # Zona 1 (Atas): MAKALAH -> Judul -> Mata Kuliah/Subjudul -> Dosen
    c_top = tbl_cov.cell(0, 0)
    t_idx = 0
    _add_cov_p(c_top, t_idx, "MAKALAH", 16, bold=True, sa=14, ls=1.2)
    t_idx += 1
    has_meta_below = bool(meta["subtitle"] or meta["course"] or meta["lecturer"])
    _add_cov_p(c_top, t_idx, meta["title"], 14, bold=True, sa=10 if has_meta_below else 0, ls=1.35)
    t_idx += 1
    if meta["subtitle"]:
        _add_cov_p(c_top, t_idx, meta["subtitle"], 12, bold=False, sa=4 if meta["lecturer"] else 0, ls=1.25)
        t_idx += 1
    elif meta["course"]:
        _add_cov_p(c_top, t_idx, f"Mata Kuliah: {meta['course']}", 12, bold=False, sa=3 if meta["lecturer"] else 0, ls=1.25)
        t_idx += 1
    if meta["lecturer"]:
        _add_cov_p(c_top, t_idx, f"Dosen Pengampu: {meta['lecturer']}", 12, bold=False, sa=0, ls=1.25)

    # Zona 2 (Tengah): Logo -> Disusun Oleh -> Penulis/NIM
    c_mid = tbl_cov.cell(1, 0)
    m_idx = 0
    if has_logo and logo_w_cm > 0 and logo_h_cm > 0:
        p_logo = c_mid.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.first_line_indent = Cm(0)
        p_logo.paragraph_format.line_spacing = 1.0
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(20)
        p_logo.add_run().add_picture(logo_path, width=Cm(logo_w_cm), height=Cm(logo_h_cm))
        m_idx = 1
    _add_cov_p(c_mid, m_idx, "Disusun Oleh:", 12, bold=True, sa=5, ls=1.25)
    if len(authors) > 1:
        tbl_auth = c_mid.add_table(rows=len(authors), cols=2)
        tbl_auth.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_borderless_table(tbl_auth)
        for idx, auth in enumerate(authors):
            c0 = tbl_auth.cell(idx, 0)
            c1 = tbl_auth.cell(idx, 1)
            c0.width = Cm(6.5)
            c1.width = Cm(4.0)
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p0.paragraph_format.first_line_indent = Cm(0)
            p0.paragraph_format.space_after = Pt(2)
            p0.paragraph_format.line_spacing = 1.15
            add_formatted_runs(p0, auth.get("name", ""), font_size_pt=12, default_bold=True)

            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p1.paragraph_format.first_line_indent = Cm(0)
            p1.paragraph_format.space_after = Pt(2)
            p1.paragraph_format.line_spacing = 1.15
            id_str = (auth.get("id") or "").strip()
            add_formatted_runs(p1, f"NIM. {id_str}" if id_str else "", font_size_pt=12, default_bold=True)
    else:
        auth = authors[0] if authors else {"name": "", "id": ""}
        id_str = (auth.get("id") or "").strip()
        _add_cov_p(c_mid, m_idx + 1, auth.get("name", ""), 12, bold=True, sa=2 if id_str else 0, ls=1.25)
        if id_str:
            _add_cov_p(c_mid, m_idx + 2, f"NIM. {id_str}", 12, bold=True, sa=0, ls=1.25)

    # Zona 3 (Bawah Rapat Margin): Program Studi -> Fakultas -> Institusi -> Kota -> Tahun Akademik
    c_bot = tbl_cov.cell(2, 0)
    for idx_i, iline in enumerate(meta["inst_lines"]):
        is_last_i = idx_i == len(meta["inst_lines"]) - 1
        _add_cov_p(c_bot, idx_i, iline, 13, bold=True, sa=0 if is_last_i else 2, ls=1.25)

    p_cov_tail = doc.add_paragraph()
    p_cov_tail.paragraph_format.first_line_indent = Cm(0)
    p_cov_tail.paragraph_format.space_before = Pt(0)
    p_cov_tail.paragraph_format.space_after = Pt(0)
    p_cov_tail.paragraph_format.line_spacing = 1.0
    r_tail = p_cov_tail.add_run("")
    r_tail.font.size = Pt(1)

    # SECTION 1: BAGIAN AWAL (Romawi Kecil Bawah Tengah)
    sec_prelim = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_preliminary_pagination(sec_prelim, restart_at_one=True)

    # ABSTRAK (Opsional)
    if abstract_blocks:
        p_abs = doc.add_paragraph(style="Heading 1")
        p_abs.paragraph_format.space_after = Pt(14)
        add_formatted_runs(p_abs, "ABSTRAK", font_size_pt=12, default_bold=True)
        for b in abstract_blocks:
            if b["type"] == "paragraph":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.line_spacing = cfg["typography"]["single_spacing"]
                p.paragraph_format.space_after = Pt(6)
                add_formatted_runs(p, b["text"], font_size_pt=12)
        doc.add_page_break()

    # KATA PENGANTAR
    p_kp = doc.add_paragraph(style="Heading 1")
    add_formatted_runs(p_kp, "KATA PENGANTAR", font_size_pt=12, default_bold=True)

    _render_docx_blocks(doc, preface_blocks, base_dir)

    has_manual_sign = any(
        b["type"] == "paragraph" and re.search(r"\bpenulis\s*$", b["text"].strip(), re.IGNORECASE)
        for b in preface_blocks[-2:]
    )
    if not has_manual_sign:
        city_date = blueprint.get("preface_date", f"{blueprint.get('city', 'Jakarta')}, {blueprint.get('year', '2026')}")
        p_date = doc.add_paragraph()
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_date.paragraph_format.first_line_indent = Cm(0)
        p_date.paragraph_format.space_before = Pt(18)
        p_date.paragraph_format.space_after = Pt(36)
        add_formatted_runs(p_date, city_date, font_size_pt=12)

        p_sign = doc.add_paragraph()
        p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_sign.paragraph_format.first_line_indent = Cm(0)
        add_formatted_runs(p_sign, "Penulis", font_size_pt=12, default_bold=True)

    # DAFTAR ISI (Native Word Table of Contents: TOC \o "1-3" \h \z \u)
    doc.add_page_break()
    p_toc = doc.add_paragraph()
    p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc.paragraph_format.first_line_indent = Cm(0)
    p_toc.paragraph_format.space_after = Pt(18)
    p_toc.paragraph_format.keep_with_next = True
    add_formatted_runs(p_toc, "DAFTAR ISI", font_size_pt=12, default_bold=True)

    raw_toc = toc_data.get("toc", []) if toc_data else []
    word_toc_entries = [item for item in raw_toc if item.get("label", "").upper() != "DAFTAR ISI"]
    if not word_toc_entries:
        word_toc_entries = [
            {"label": "KATA PENGANTAR", "page": "i", "indent_cm": 0.0, "bold": True},
        ]
    add_native_word_toc(doc, 'TOC \\o "1-3" \\h \\z \\u', word_toc_entries, gallery_name="Table of Contents")

    # DAFTAR TABEL (Native Word TOC untuk style Caption Tabel)
    lot_items = toc_data.get("lot", []) if toc_data else []
    if lot_items:
        doc.add_page_break()
        p_lot = doc.add_paragraph(style="Heading 1")
        add_formatted_runs(p_lot, "DAFTAR TABEL", font_size_pt=12, default_bold=True)
        add_native_word_toc(doc, 'TOC \\h \\z \\t "Caption Tabel,1"', lot_items, gallery_name="Table of Figures")

    # DAFTAR GAMBAR (Native Word TOC untuk style Caption Gambar)
    lof_items = toc_data.get("lof", []) if toc_data else []
    if lof_items:
        doc.add_page_break()
        p_lof = doc.add_paragraph(style="Heading 1")
        add_formatted_runs(p_lof, "DAFTAR GAMBAR", font_size_pt=12, default_bold=True)
        add_native_word_toc(doc, 'TOC \\h \\z \\t "Caption Gambar,1"', lof_items, gallery_name="Table of Figures")

    # SECTION 2..N: BAB I s.d. BAB III
    for ch_idx, (ch_meta, blocks) in enumerate(chapter_blocks):
        sec_ch = doc.add_section(WD_SECTION.NEW_PAGE)
        setup_body_chapter_pagination(sec_ch, restart_at_one=(ch_idx == 0))

        # Satu paragraf Heading 1 dengan soft line break agar tampil 2 baris di halaman & 1 baris di TOC Word
        p_bab = doc.add_paragraph(style="Heading 1")
        ch_num_str = ch_meta.get("number", f"BAB {ch_idx + 1}").upper()
        ch_ttl_str = ch_meta.get("title", "").upper()
        add_formatted_runs(p_bab, ch_num_str, font_size_pt=12, default_bold=True)
        if ch_ttl_str:
            r_br = p_bab.add_run()
            r_br.add_break()
            add_formatted_runs(p_bab, ch_ttl_str, font_size_pt=12, default_bold=True)

        _render_docx_blocks(doc, blocks, base_dir)

    # SECTION TERAKHIR: DAFTAR PUSTAKA
    sec_ref = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_body_chapter_pagination(sec_ref, restart_at_one=False)

    p_dp = doc.add_paragraph(style="Heading 1")
    add_formatted_runs(p_dp, "DAFTAR PUSTAKA", font_size_pt=12, default_bold=True)

    hang_cm = cfg["typography"]["hanging_indent_cm"]
    for ref_item in get_sorted_references(blueprint):
        p_r = doc.add_paragraph()
        p_r.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf_r = p_r.paragraph_format
        pf_r.left_indent = Cm(hang_cm)
        pf_r.first_line_indent = Cm(-hang_cm)
        pf_r.space_after = Pt(4)
        pf_r.line_spacing = cfg["typography"]["single_spacing"]
        pf_r.keep_together = True
        pf_r.widow_control = True
        add_formatted_runs(p_r, ref_item, font_size_pt=12)

    os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)
    doc.save(output_docx_path)


# ==============================================================================
# PDF GENERATION ENGINE (reportlab Two-Pass Unicode)
# ==============================================================================

def register_pdf_fonts() -> tuple[str, str, str, str]:
    """Daftarkan Times New Roman TTF Windows untuk dukungan Unicode penuh di PDF."""
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    candidates = {
        "TNR": os.path.join(fonts_dir, "times.ttf"),
        "TNR-Bold": os.path.join(fonts_dir, "timesbd.ttf"),
        "TNR-Italic": os.path.join(fonts_dir, "timesi.ttf"),
        "TNR-BoldItalic": os.path.join(fonts_dir, "timesbi.ttf"),
    }
    if all(os.path.exists(p) for p in candidates.values()):
        for name, path in candidates.items():
            if name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(name, path))
        registerFontFamily("TNR", normal="TNR", bold="TNR-Bold", italic="TNR-Italic", boldItalic="TNR-BoldItalic")
        return "TNR", "TNR-Bold", "TNR-Italic", "TNR-BoldItalic"
    return "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"


def md_inline_to_rl_xml(text: str, default_bold: bool = False, default_italic: bool = False) -> str:
    parts = []
    for seg, is_bold, is_italic in tokenize_inline_markdown(text, default_bold, default_italic):
        esc = html.escape(seg)
        if is_bold and is_italic:
            parts.append(f"<b><i>{esc}</i></b>")
        elif is_bold:
            parts.append(f"<b>{esc}</b>")
        elif is_italic:
            parts.append(f"<i>{esc}</i>")
        else:
            parts.append(esc)
    return "".join(parts)


class TocMarker(Flowable):
    """Marker tidak terlihat untuk mencatat halaman aktual elemen pada Pass 1 PDF."""

    def __init__(self, kind: str, label: str, indent_cm: float = 0.0, bold: bool = False):
        super().__init__()
        self.kind = kind
        self.label = label
        self.indent_cm = indent_cm
        self.bold = bold
        self.width = 0
        self.height = 0

    def draw(self):
        pass


class MakalahPDFDocTemplate(BaseDocTemplate):
    """DocTemplate Two-Pass yang mencatat nomor halaman aktual Romawi & Arab."""

    def __init__(self, filename_or_buf, **kw):
        super().__init__(filename_or_buf, **kw)
        self.prelim_counter = 0
        self.body_counter = 0
        self.recorded_toc: list[dict] = []
        self.recorded_lot: list[dict] = []
        self.recorded_lof: list[dict] = []

    def handle_pageBegin(self):
        # Increment counter SEBELUM super().handle_pageBegin() karena onPage() dipanggil di dalamnya
        if self.pageTemplate.id == "prelim":
            self.prelim_counter += 1
        elif self.pageTemplate.id in ("chapter_first", "body_cont"):
            self.body_counter += 1
        super().handle_pageBegin()

    def afterFlowable(self, flowable):
        if isinstance(flowable, TocMarker):
            if self.pageTemplate.id == "prelim":
                page_str = to_roman_lower(self.prelim_counter)
            else:
                page_str = str(self.body_counter)
            entry = {
                "label": flowable.label,
                "page": page_str,
                "indent_cm": flowable.indent_cm,
                "bold": flowable.bold,
            }
            if flowable.kind == "toc":
                self.recorded_toc.append(entry)
            elif flowable.kind == "lot":
                self.recorded_lot.append(entry)
            elif flowable.kind == "lof":
                self.recorded_lof.append(entry)


def _compute_pdf_col_widths(rows: list[list[str]], total_width_pts: float) -> list[float]:
    return [w * cm for w in _compute_col_widths_cm(rows, total_width_pts / cm)]


class DottedTocRow(Flowable):
    """Baris Daftar Isi PDF dengan titik-titik leader klasik hingga nomor halaman rata kanan."""

    def __init__(self, item: dict, frame_w: float, f_reg: str, f_bold: str, base_style):
        super().__init__()
        self.item = item
        self.frame_w = frame_w
        self.f_reg = f_reg
        self.f_bold = f_bold
        self.font_used = f_bold if item["bold"] else f_reg
        self.indent_pts = item["indent_cm"] * cm
        self.avail_label_w = max(4.0 * cm, frame_w - self.indent_pts - 1.3 * cm)
        st_lbl = ParagraphStyle(
            f"TOCL_{item['bold']}_{item['indent_cm']}",
            parent=base_style,
            fontName=self.font_used,
            fontSize=12,
            leading=17,
            alignment=TA_LEFT,
        )
        self.para = Paragraph(md_inline_to_rl_xml(item["label"], default_bold=item["bold"]), st_lbl)
        self.last_line_end_x = self.indent_pts

    def wrap(self, availWidth, availHeight):
        pw, ph = self.para.wrap(self.avail_label_w, availHeight)
        self.width = self.frame_w
        self.height = ph + 3
        try:
            last_line = self.para.blPara.lines[-1]
            extra = last_line.extraSpace if hasattr(last_line, "extraSpace") else last_line[0]
            self.last_line_end_x = self.indent_pts + max(0.0, self.avail_label_w - extra)
        except Exception:
            plain = _strip_all_md_markers(self.item["label"])
            self.last_line_end_x = min(self.indent_pts + pdfmetrics.stringWidth(plain, self.font_used, 12), self.frame_w - 1.5 * cm)
        return self.width, self.height

    def draw(self):
        self.para.drawOn(self.canv, self.indent_pts, 3)
        page_str = self.item["page"]
        pg_w = pdfmetrics.stringWidth(page_str, self.font_used, 12)
        dot_start = self.last_line_end_x + 5
        dot_end = self.frame_w - pg_w - 6
        self.canv.saveState()
        self.canv.setFont(self.f_reg, 11)
        x = dot_start
        while x < dot_end:
            self.canv.drawString(x, 6.5, ".")
            x += 4.2
        self.canv.setFont(self.font_used, 12)
        self.canv.drawRightString(self.frame_w, 6.5, page_str)
        self.canv.restoreState()


def _render_pdf_toc_table(entries: list[dict], frame_w: float, f_reg: str, f_bold: str, styles) -> list:
    if not entries:
        return []
    return [DottedTocRow(item, frame_w, f_reg, f_bold, styles["Normal"]) for item in entries]


def _build_pdf_story(
    blueprint: dict,
    abstract_blocks: list[dict],
    preface_blocks: list[dict],
    chapter_blocks: list[tuple[dict, list[dict]]],
    logo_path: str,
    base_dir: str,
    frame_w: float,
    fonts: tuple[str, str, str, str],
    toc_snapshot: dict | None = None,
) -> list:
    f_reg, f_bold, f_ital, _ = fonts
    styles = getSampleStyleSheet()

    st_cover_title = ParagraphStyle("CoverTitle", parent=styles["Normal"], fontName=f_bold, fontSize=14, leading=21, alignment=TA_CENTER, spaceAfter=10)
    st_cover_text = ParagraphStyle("CoverText", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=18, alignment=TA_CENTER, spaceAfter=4)
    st_cover_bold = ParagraphStyle("CoverBold", parent=st_cover_text, fontName=f_bold)
    st_chapter_title = ParagraphStyle("ChapterTitle", parent=styles["Normal"], fontName=f_bold, fontSize=12, leading=18, alignment=TA_CENTER, spaceAfter=14, keepWithNext=True)
    st_subheading = ParagraphStyle("SubHeading", parent=styles["Normal"], fontName=f_bold, fontSize=12, leading=18, alignment=TA_LEFT, spaceBefore=10, spaceAfter=4, keepWithNext=True)
    st_body = ParagraphStyle("BodyAcad", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=18, alignment=TA_JUSTIFY, firstLineIndent=cfg["typography"]["first_line_indent_cm"] * cm, spaceAfter=6)
    st_abs = ParagraphStyle("AbsAcad", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=13.5, alignment=TA_JUSTIFY, firstLineIndent=0, spaceAfter=6)
    st_blockquote = ParagraphStyle("QuoteAcad", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=13.5, alignment=TA_JUSTIFY, leftIndent=cfg["typography"]["blockquote_indent_cm"] * cm, firstLineIndent=0, spaceBefore=6, spaceAfter=6)
    st_ref = ParagraphStyle("RefAcad", parent=styles["Normal"], fontName=f_reg, fontSize=12, leading=13, alignment=TA_JUSTIFY, leftIndent=cfg["typography"]["hanging_indent_cm"] * cm, firstLineIndent=-cfg["typography"]["hanging_indent_cm"] * cm, spaceAfter=4)
    st_caption = ParagraphStyle("CapAcad", parent=styles["Normal"], fontName=f_bold, fontSize=11, leading=14, alignment=TA_LEFT, spaceBefore=8, spaceAfter=4, keepWithNext=True)
    st_source = ParagraphStyle("SrcAcad", parent=styles["Normal"], fontName=f_ital, fontSize=10, leading=13, alignment=TA_LEFT, spaceBefore=3, spaceAfter=8)
    st_right = ParagraphStyle("RightSign", parent=st_cover_text, alignment=TA_RIGHT)

    def _make_simple_flow(blk: dict):
        bt = blk["type"]
        if bt == "paragraph":
            return Paragraph(md_inline_to_rl_xml(blk["text"]), st_body)
        if bt == "blockquote":
            return Paragraph(md_inline_to_rl_xml(blk["text"]), st_blockquote)
        extra = 0.5 * blk.get("indent_level", 0)
        st_li = ParagraphStyle(
            f"ListAcad_{extra}",
            parent=styles["Normal"],
            fontName=f_reg,
            fontSize=12,
            leading=18,
            alignment=TA_JUSTIFY,
            leftIndent=(0.75 + extra) * cm,
            firstLineIndent=-0.75 * cm,
            spaceAfter=3,
        )
        return Paragraph(f"{html.escape(blk['marker'])} {md_inline_to_rl_xml(blk['text'])}", st_li)

    def render_blocks_pdf(blks: list[dict], target_story: list, track_items: bool = False) -> None:
        i = 0
        while i < len(blks):
            b = blks[i]
            btype = b["type"]
            if btype == "heading":
                h_grp: list = []
                while i < len(blks) and blks[i]["type"] == "heading":
                    hb = blks[i]
                    if track_items:
                        indent = 0.75 if hb["level"] == 2 else 1.5
                        h_grp.append(TocMarker("toc", hb["text"], indent_cm=indent, bold=False))
                    h_grp.append(Paragraph(md_inline_to_rl_xml(hb["text"], default_bold=True), st_subheading))
                    i += 1
                if i < len(blks) and blks[i]["type"] in ("paragraph", "blockquote", "list_item"):
                    h_grp.append(_make_simple_flow(blks[i]))
                    i += 1
                target_story.append(KeepTogether(h_grp))
                continue
            elif btype in ("paragraph", "blockquote", "list_item"):
                target_story.append(_make_simple_flow(b))
            elif btype == "table":
                t_elems = []
                if track_items:
                    t_elems.append(TocMarker("lot", b["caption"]))
                t_elems.append(Paragraph(md_inline_to_rl_xml(b["caption"], default_bold=True), st_caption))
                rows = b.get("rows", [])
                if rows:
                    num_cols = max(len(r) for r in rows)
                    norm_rows = [r + [""] * (num_cols - len(r)) for r in rows]
                    col_widths = _compute_pdf_col_widths(norm_rows, frame_w)
                    rl_rows = []
                    for r_idx, r in enumerate(norm_rows):
                        c_style = ParagraphStyle(
                            f"TCell_{r_idx}",
                            parent=styles["Normal"],
                            fontName=f_bold if r_idx == 0 else f_reg,
                            fontSize=10.5,
                            leading=13,
                            alignment=TA_CENTER if r_idx == 0 else TA_LEFT,
                        )
                        rl_rows.append([Paragraph(md_inline_to_rl_xml(c, default_bold=(r_idx == 0)), c_style) for c in r])
                    tbl = Table(rl_rows, colWidths=col_widths, repeatRows=1)
                    if cfg["table"].get("style", "word_grid") == "word_grid":
                        pdf_tbl_cmds = [
                            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(f"#{cfg['table'].get('header_bg', 'F2F2F2')}")),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                        ]
                    else:
                        pdf_tbl_cmds = [
                            ("LINEABOVE", (0, 0), (-1, 0), 1.2, colors.black),
                            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
                            ("LINEBELOW", (0, -1), (-1, -1), 1.2, colors.black),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    tbl.setStyle(TableStyle(pdf_tbl_cmds))
                    t_elems.append(tbl)
                if b.get("source"):
                    t_elems.append(Paragraph(md_inline_to_rl_xml(b["source"], default_italic=True), st_source))
                target_story.append(KeepTogether(t_elems))
            elif btype == "figure":
                f_elems = []
                if track_items:
                    f_elems.append(TocMarker("lof", b["caption"]))
                img_rel = b.get("path", "")
                img_abs = img_rel if os.path.isabs(img_rel) else os.path.join(base_dir, img_rel)
                if os.path.exists(img_abs):
                    im = RLImage(img_abs, width=10.0 * cm, height=7.0 * cm, kind="proportional", mask="auto")
                    im.hAlign = "CENTER"
                    f_elems.append(im)
                st_fcap = ParagraphStyle("FigCap", parent=st_caption, alignment=TA_CENTER)
                st_fsrc = ParagraphStyle("FigSrc", parent=st_source, alignment=TA_CENTER)
                f_elems.append(Paragraph(md_inline_to_rl_xml(b["caption"], default_bold=True), st_fcap))
                if b.get("source"):
                    f_elems.append(Paragraph(md_inline_to_rl_xml(b["source"], default_italic=True), st_fsrc))
                target_story.append(KeepTogether(f_elems))
            i += 1

    story: list = []
    cov_spec = _build_cover_zone_spec(blueprint, logo_path, fonts=fonts)

    # 1. COVER PAGE (Grid 3-Zona Penuh: Atas-Tengah-Bawah Rapat Margin)
    cov_tbl = Table(
        [[cov_spec["top_flows"]], [cov_spec["mid_flows"]], [cov_spec["bot_flows"]]],
        colWidths=[cov_spec["cov_frame_w"]],
        rowHeights=list(cov_spec["row_heights"]),
    )
    cov_tbl.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (0, 0), "TOP"),
            ("VALIGN", (0, 1), (0, 1), "MIDDLE"),
            ("VALIGN", (0, 2), (0, 2), "BOTTOM"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ])
    )
    story.append(cov_tbl)

    # 2. PRELIMINARY
    story.append(NextPageTemplate("prelim"))
    story.append(PageBreak())

    if abstract_blocks:
        story.append(TocMarker("toc", "ABSTRAK", bold=True))
        story.append(Paragraph("ABSTRAK", st_chapter_title))
        for b in abstract_blocks:
            if b["type"] == "paragraph":
                story.append(Paragraph(md_inline_to_rl_xml(b["text"]), st_abs))
        story.append(PageBreak())

    story.append(TocMarker("toc", "KATA PENGANTAR", bold=True))
    story.append(Paragraph("KATA PENGANTAR", st_chapter_title))
    render_blocks_pdf(preface_blocks, story, track_items=False)

    has_manual_sign = any(
        b["type"] == "paragraph" and re.search(r"\bpenulis\s*$", b["text"].strip(), re.IGNORECASE)
        for b in preface_blocks[-2:]
    )
    if not has_manual_sign:
        story.append(Spacer(1, 0.7 * cm))
        city_date = blueprint.get("preface_date", f"{blueprint.get('city', 'Jakarta')}, {blueprint.get('year', '2026')}")
        story.append(Paragraph(html.escape(city_date), st_right))
        story.append(Spacer(1, 0.9 * cm))
        story.append(Paragraph("<b>Penulis</b>", st_right))

    # DAFTAR ISI
    story.append(PageBreak())
    story.append(TocMarker("toc", "DAFTAR ISI", bold=True))
    story.append(Paragraph("DAFTAR ISI", st_chapter_title))

    all_tables = [b["caption"] for _, blks in chapter_blocks for b in blks if b["type"] == "table"]
    all_figures = [b["caption"] for _, blks in chapter_blocks for b in blks if b["type"] == "figure"]

    if toc_snapshot and toc_snapshot.get("toc"):
        story.extend(_render_pdf_toc_table(toc_snapshot["toc"], frame_w, f_reg, f_bold, styles))
    else:
        placeholder_toc = []
        if abstract_blocks:
            placeholder_toc.append({"label": "ABSTRAK", "page": "i", "indent_cm": 0.0, "bold": True})
        placeholder_toc.append({"label": "KATA PENGANTAR", "page": "i", "indent_cm": 0.0, "bold": True})
        placeholder_toc.append({"label": "DAFTAR ISI", "page": "ii", "indent_cm": 0.0, "bold": True})
        if all_tables:
            placeholder_toc.append({"label": "DAFTAR TABEL", "page": "iii", "indent_cm": 0.0, "bold": True})
        if all_figures:
            placeholder_toc.append({"label": "DAFTAR GAMBAR", "page": "iv", "indent_cm": 0.0, "bold": True})
        for ch_meta, blks in chapter_blocks:
            ch_full = f"{ch_meta.get('number', 'BAB').upper()} {ch_meta.get('title', '').upper()}".strip()
            placeholder_toc.append({"label": ch_full, "page": "1", "indent_cm": 0.0, "bold": True})
            for b in blks:
                if b["type"] == "heading":
                    placeholder_toc.append({"label": b["text"], "page": "1", "indent_cm": 0.75 if b["level"] == 2 else 1.5, "bold": False})
        placeholder_toc.append({"label": "DAFTAR PUSTAKA", "page": "1", "indent_cm": 0.0, "bold": True})
        story.extend(_render_pdf_toc_table(placeholder_toc, frame_w, f_reg, f_bold, styles))

    if all_tables:
        story.append(PageBreak())
        story.append(TocMarker("toc", "DAFTAR TABEL", bold=True))
        story.append(Paragraph("DAFTAR TABEL", st_chapter_title))
        lot_entries = toc_snapshot["lot"] if toc_snapshot and toc_snapshot.get("lot") else [{"label": c, "page": "1", "indent_cm": 0.0, "bold": False} for c in all_tables]
        story.extend(_render_pdf_toc_table(lot_entries, frame_w, f_reg, f_bold, styles))

    if all_figures:
        story.append(PageBreak())
        story.append(TocMarker("toc", "DAFTAR GAMBAR", bold=True))
        story.append(Paragraph("DAFTAR GAMBAR", st_chapter_title))
        lof_entries = toc_snapshot["lof"] if toc_snapshot and toc_snapshot.get("lof") else [{"label": c, "page": "1", "indent_cm": 0.0, "bold": False} for c in all_figures]
        story.extend(_render_pdf_toc_table(lof_entries, frame_w, f_reg, f_bold, styles))

    # 3. BAB I s.d. BAB III
    for ch_idx, (ch_meta, blocks) in enumerate(chapter_blocks):
        story.append(NextPageTemplate("chapter_first"))
        story.append(PageBreak())
        story.append(NextPageTemplate("body_cont"))
        ch_num = ch_meta.get("number", f"BAB {ch_idx + 1}").upper()
        ch_ttl = ch_meta.get("title", "").upper()
        story.append(TocMarker("toc", f"{ch_num} {ch_ttl}".strip(), bold=True))
        story.append(Paragraph(f"{html.escape(ch_num)}<br/>{html.escape(ch_ttl)}", st_chapter_title))
        render_blocks_pdf(blocks, story, track_items=True)

    # 4. DAFTAR PUSTAKA
    story.append(NextPageTemplate("chapter_first"))
    story.append(PageBreak())
    story.append(NextPageTemplate("body_cont"))
    story.append(TocMarker("toc", "DAFTAR PUSTAKA", bold=True))
    story.append(Paragraph("DAFTAR PUSTAKA", st_chapter_title))
    for ref_item in get_sorted_references(blueprint):
        story.append(KeepTogether([Paragraph(md_inline_to_rl_xml(ref_item), st_ref)]))

    return story


def build_pdf(
    blueprint: dict,
    abstract_blocks: list[dict],
    preface_blocks: list[dict],
    chapter_blocks: list[tuple[dict, list[dict]]],
    logo_path: str,
    base_dir: str,
    output_pdf_path: str,
) -> dict:
    """Bangun PDF dalam 2 pass dan kembalikan peta halaman aktual untuk sinkronisasi TOC."""
    os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
    fonts = register_pdf_fonts()
    f_reg = fonts[0]

    left_m = cfg["page"]["margin_left_cm"] * cm
    top_m = cfg["page"]["margin_top_cm"] * cm
    right_m = cfg["page"]["margin_right_cm"] * cm
    bottom_m = cfg["page"]["margin_bottom_cm"] * cm
    frame_w = A4[0] - left_m - right_m
    frame_h = A4[1] - top_m - bottom_m

    def make_doc(target):
        doc_obj = MakalahPDFDocTemplate(
            target,
            pagesize=A4,
            leftMargin=left_m,
            rightMargin=right_m,
            topMargin=top_m,
            bottomMargin=bottom_m,
        )

        def on_cover_page(canvas, d):
            pass

        def on_prelim_page(canvas, d):
            canvas.saveState()
            canvas.setFont(f_reg, 12)
            canvas.drawCentredString(A4[0] / 2.0, 1.5 * cm, to_roman_lower(d.prelim_counter))
            canvas.restoreState()

        def on_chapter_first_page(canvas, d):
            canvas.saveState()
            canvas.setFont(f_reg, 12)
            canvas.drawCentredString(A4[0] / 2.0, 1.5 * cm, str(d.body_counter))
            canvas.restoreState()

        def on_body_cont_page(canvas, d):
            canvas.saveState()
            canvas.setFont(f_reg, 12)
            if cfg["pagination"].get("body_other_page_pos", "bottom_center") == "bottom_center":
                canvas.drawCentredString(A4[0] / 2.0, 1.5 * cm, str(d.body_counter))
            else:
                canvas.drawRightString(A4[0] - right_m, A4[1] - 1.5 * cm, str(d.body_counter))
            canvas.restoreState()

        cov_lr_m = 3.0 * cm
        cov_frame_w = A4[0] - 2 * cov_lr_m
        doc_obj.addPageTemplates([
            PageTemplate(
                id="cover",
                frames=[Frame(cov_lr_m, bottom_m, cov_frame_w, frame_h, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="f_cov")],
                onPage=on_cover_page,
            ),
            PageTemplate(id="prelim", frames=[Frame(left_m, bottom_m, frame_w, frame_h, id="f_pre")], onPage=on_prelim_page),
            PageTemplate(id="chapter_first", frames=[Frame(left_m, bottom_m, frame_w, frame_h, id="f_ch1")], onPage=on_chapter_first_page),
            PageTemplate(id="body_cont", frames=[Frame(left_m, bottom_m, frame_w, frame_h, id="f_bdy")], onPage=on_body_cont_page),
        ])
        return doc_obj

    # Pass 1: Hitung nomor halaman aktual ke memory buffer
    buf = io.BytesIO()
    doc_p1 = make_doc(buf)
    story_p1 = _build_pdf_story(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, frame_w, fonts, toc_snapshot=None)
    doc_p1.build(story_p1)
    toc_snapshot = {
        "toc": doc_p1.recorded_toc,
        "lot": doc_p1.recorded_lot,
        "lof": doc_p1.recorded_lof,
    }

    # Pass 2: Render PDF final dengan Daftar Isi, Daftar Tabel, dan Daftar Gambar bernomor akurat
    doc_p2 = make_doc(output_pdf_path)
    story_p2 = _build_pdf_story(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, frame_w, fonts, toc_snapshot=toc_snapshot)
    doc_p2.build(story_p2)
    return toc_snapshot


# ==============================================================================
# WORKSPACE COMPILER & MANDATORY SMOKE TEST
# ==============================================================================

def compile_workspace(
    workspace_dir: str,
    output_dir: str | None = None,
    margin_override: str | None = None,
) -> dict:
    base_dir = os.path.abspath(workspace_dir)
    bp_path = os.path.join(base_dir, "blueprint.json")
    if not os.path.exists(bp_path):
        raise FileNotFoundError(f"blueprint.json tidak ditemukan di: {bp_path}")

    with open(bp_path, "r", encoding="utf-8") as f:
        blueprint = json.load(f)

    if blueprint.get("margin_preset"):
        apply_margin_preset(blueprint["margin_preset"])
    if margin_override:
        apply_margin_preset(margin_override)

    skill_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    default_skill_logo = os.path.join(skill_root, "assets", "logo_default.png")
    try:
        ensure_default_logo(default_skill_logo)
    except OSError:
        pass

    ws_assets_dir = os.path.join(base_dir, "assets")
    os.makedirs(ws_assets_dir, exist_ok=True)
    ws_default_logo = os.path.join(ws_assets_dir, "logo_default.png")
    if os.path.exists(default_skill_logo):
        if not os.path.exists(ws_default_logo):
            shutil.copyfile(default_skill_logo, ws_default_logo)
    else:
        ensure_default_logo(ws_default_logo)

    logo_path = resolve_cover_logo(blueprint, base_dir, ws_default_logo)

    abstract_blocks = []
    if blueprint.get("include_abstract"):
        abs_path = os.path.join(base_dir, blueprint.get("abstract_file", "drafts/abstrak.md"))
        if os.path.exists(abs_path):
            with open(abs_path, "r", encoding="utf-8") as af:
                abstract_blocks = parse_markdown_blocks(af.read())

    preface_rel = blueprint.get("preface_file", "drafts/kata_pengantar.md")
    preface_path = os.path.join(base_dir, preface_rel)
    preface_md = ""
    if os.path.exists(preface_path):
        with open(preface_path, "r", encoding="utf-8") as pf:
            preface_md = pf.read()
    preface_blocks = parse_markdown_blocks(preface_md)

    chapter_blocks = []
    chapter_map_for_audit = {}
    missing_chapters = []
    for idx, ch in enumerate(blueprint.get("chapters", []), 1):
        ch_rel = ch.get("file", f"drafts/bab_{idx:02d}.md")
        ch_path = os.path.join(base_dir, ch_rel)
        ch_num = ch.get("number", f"BAB {idx}")
        if not os.path.exists(ch_path):
            missing_chapters.append(f"[{ch_num}] File draf bab tidak ditemukan: '{ch_rel}'.")
            ch_md = ""
        else:
            with open(ch_path, "r", encoding="utf-8") as cf:
                ch_md = cf.read()
            if not ch_md.strip():
                missing_chapters.append(f"[{ch_num}] File draf bab masih kosong: '{ch_rel}'.")
        parsed = parse_markdown_blocks(ch_md)
        chapter_blocks.append((ch, parsed))
        chapter_map_for_audit[ch_num] = parsed

    audit_report = audit_makalah(blueprint, chapter_map_for_audit, base_dir, missing_chapters)

    out_dir = os.path.abspath(output_dir) if output_dir else os.path.join(base_dir, "output")
    os.makedirs(out_dir, exist_ok=True)

    slug = re.sub(r"[^a-zA-Z0-9]+", "_", blueprint.get("title", "Makalah_Akademik")).strip("_")[:50]
    if not slug:
        slug = "Makalah_Akademik"
    docx_out = os.path.join(out_dir, f"{slug}.docx")
    pdf_out = os.path.join(out_dir, f"{slug}.pdf")

    # Bangun PDF terlebih dahulu (Two-Pass) untuk mendapatkan nomor halaman TOC yang akurat bagi DOCX & PDF
    try:
        toc_snapshot = build_pdf(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, pdf_out)
    except PermissionError:
        rev = 1
        while True:
            cand_pdf = os.path.join(out_dir, f"{slug}_revisi_{rev}.pdf")
            try:
                toc_snapshot = build_pdf(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, cand_pdf)
                pdf_out = cand_pdf
                break
            except PermissionError:
                rev += 1

    try:
        build_docx(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, docx_out, toc_data=toc_snapshot)
    except PermissionError:
        rev = 1
        while True:
            cand_docx = os.path.join(out_dir, f"{slug}_revisi_{rev}.docx")
            try:
                build_docx(blueprint, abstract_blocks, preface_blocks, chapter_blocks, logo_path, base_dir, cand_docx, toc_data=toc_snapshot)
                docx_out = cand_docx
                break
            except PermissionError:
                rev += 1

    return {
        "docx_path": docx_out,
        "pdf_path": pdf_out,
        "logo_path": logo_path,
        "audit": audit_report,
        "toc": toc_snapshot,
    }


def run_smoke_test() -> int:
    """Mandatory Smoke Test untuk memverifikasi integritas DOCX, XML, Logo, Figure, Tabel, dan PDF Two-Pass."""
    print("=== RUNNING MANDATORY SMOKE TEST (penyusun-makalah) ===")
    tmp_dir = tempfile.mkdtemp(prefix="makalah_smoke_")
    try:
        drafts_dir = os.path.join(tmp_dir, "drafts")
        assets_dir = os.path.join(tmp_dir, "assets")
        os.makedirs(drafts_dir, exist_ok=True)
        os.makedirs(assets_dir, exist_ok=True)

        # Buat gambar diagram uji sederhana (RGBA transparan)
        fig_path = os.path.join(assets_dir, "arsitektur_ai.png")
        fig_img = Image.new("RGBA", (600, 320), (245, 248, 255, 255))
        d = ImageDraw.Draw(fig_img)
        d.rectangle((30, 60, 200, 240), fill=(25, 55, 105, 255), outline=(0, 0, 0, 255), width=3)
        d.rectangle((240, 60, 410, 240), fill=(45, 105, 165, 255), outline=(0, 0, 0, 255), width=3)
        d.rectangle((450, 60, 570, 240), fill=(197, 155, 52, 255), outline=(0, 0, 0, 255), width=3)
        fig_img.save(fig_path, "PNG")

        sample_bp = {
            "title": "OPTIMALISASI RANTAI PASOK INDUSTRI MANUFAKTUR MELALUI INTEGRASI KECERDASAN BUATAN",
            "subtitle": "",
            "course": "Manajemen Operasional",
            "lecturer": "Prof. Dr. Ir. Hendra Wijaya, M.T., IPU.",
            "authors": [
                {"name": "Budi Santoso", "id": "2301010045"},
                {"name": "Siti Aminah", "id": "2301010052"},
            ],
            "department": "Program Studi Teknik Industri",
            "faculty": "Fakultas Teknik",
            "institution": "Universitas Nusantara",
            "city": "Jakarta",
            "year": "2026",
            "academic_year": "2025/2026",
            "logo_path": "",
            "numbering_style": "decimal",
            "margin_preset": "4-3-3-3",
            "preface_date": "Jakarta, November 2026",
            "preface_file": "drafts/kata_pengantar.md",
            "problem_statements": [
                "Bagaimana arsitektur penerapan kecerdasan buatan dalam optimalisasi rantai pasok industri manufaktur?",
                "Bagaimana dampak integrasi algoritma prediktif terhadap efisiensi biaya logistik perusahaan?",
            ],
            "chapters": [
                {"number": "BAB I", "title": "PENDAHULUAN", "file": "drafts/bab_01.md", "status": "completed"},
                {"number": "BAB II", "title": "PEMBAHASAN", "file": "drafts/bab_02.md", "status": "completed"},
                {"number": "BAB III", "title": "PENUTUP", "file": "drafts/bab_03.md", "status": "completed"},
            ],
            "references": [
                "Anggraini, D., & Putra, M. (2023). Transformasi logistik manufaktur berbasis data. *Jurnal Sistem Industri*, 12(1), 10-22.",
                "Budiman, T. (2024). Otomatisasi gudang dan pengendalian persediaan modern. *Jurnal Teknologi Manufaktur*, 8(2), 45-58.",
                "Christopher, M. (2022). *Logistics & supply chain management* (6th ed.). Pearson Education.",
                "Dewi, S. K., & Hartono, B. (2021). Evaluasi efisiensi distribusi pada industri nasional. *Jurnal Manajemen Operasi*, 9(3), 77-89.",
                "Fauzan, R. (2024). Arsitektur analitik prediktif untuk rantai pasok tangguh. *Jurnal Informatika Terapan*, 14(1), 33-44.",
                "Gunawan, H., & Lestari, P. (2023). Peramalan permintaan berbasis jaringan saraf tiruan. *Jurnal Rekayasa Sistem*, 11(2), 101-114.",
                "Hidayat, A. (2022). Manajemen risiko pasokan pada era industri digital. *Jurnal Ekonomi dan Industri*, 7(4), 210-221.",
                "Kusuma, W., & Santoso, I. (2025). Optimalisasi rute distribusi dengan algoritma cerdas. *Jurnal Teknik Logistik*, 6(1), 15-28.",
                "Mahendra, Y. (2023). Integrasi ERP dan modul kecerdasan buatan di pabrik. *Jurnal Sistem Informasi*, 19(2), 88-99.",
                "Pratama, R. A., & Nugroho, D. S. (2023). Integrasi *machine learning* dalam peramalan permintaan rantai pasok manufaktur. *Jurnal Teknik Industri Indonesia*, 18(2), 112-125.",
                "Wijaya, K., & Rahmawati, N. (2024). Keamanan siber dan tata kelola data rantai pasok. *Jurnal Keamanan Digital*, 5(1), 50-63.",
            ],
        }

        with open(os.path.join(tmp_dir, "blueprint.json"), "w", encoding="utf-8") as f:
            json.dump(sample_bp, f, indent=2, ensure_ascii=False)

        with open(os.path.join(drafts_dir, "kata_pengantar.md"), "w", encoding="utf-8") as f:
            f.write(
                "Puji syukur penulis panjatkan ke hadirat Tuhan Yang Maha Esa karena atas rahmat-Nya makalah ini dapat diselesaikan tepat waktu. "
                "Makalah berjudul Optimalisasi Rantai Pasok Industri Manufaktur melalui Integrasi Kecerdasan Buatan ini disusun untuk memenuhi tugas Mata Kuliah Manajemen Operasional. "
                "Penulis menyampaikan terima kasih kepada Bapak Prof. Dr. Ir. Hendra Wijaya, M.T., IPU. selaku dosen pengampu yang telah memberikan bimbingan akademis.\n\n"
                "Penulis menyadari bahwa penyusunan makalah ini masih memiliki ruang perbaikan di berbagai aspek analisis. "
                "Oleh karena itu, kritik dan saran yang bersifat membangun dari para pembaca sangat diharapkan demi penyempurnaan kajian ilmiah selanjutnya. "
                "Semoga makalah ini memberikan manfaat teoretis maupun praktis bagi pengembangan ilmu pengetahuan."
            )

        with open(os.path.join(drafts_dir, "bab_01.md"), "w", encoding="utf-8") as f:
            f.write(
                "## 1.1 Latar Belakang\n\n"
                "Transformasi digital pada sektor manufaktur global telah mengubah paradigma pengelolaan *supply chain management* secara fundamental dalam satu dekade terakhir. "
                "Ketidakpastian permintaan pasar dan fluktuasi pasokan bahan baku menuntut perusahaan manufaktur untuk mengadopsi sistem peramalan yang adaptif (Christopher, 2022). "
                "Oleh sebab itu, kajian mengenai pemanfaatan algoritma *artificial intelligence* (AI) dalam rantai pasok menjadi sangat mendesak untuk dibahas secara ilmiah.\n\n"
                "## 1.2 Rumusan Masalah\n\n"
                "Berdasarkan uraian latar belakang di atas, rumusan masalah dalam makalah ini adalah sebagai berikut:\n"
                "1. Bagaimana arsitektur penerapan *artificial intelligence* dalam optimalisasi rantai pasok industri manufaktur?\n"
                "2. Bagaimana dampak integrasi algoritma prediktif terhadap efisiensi biaya logistik perusahaan?\n\n"
                "## 1.3 Tujuan Penulisan\n\n"
                "1. Menganalisis arsitektur penerapan *artificial intelligence* dalam optimalisasi rantai pasok industri manufaktur.\n"
                "2. Mengevaluasi dampak integrasi algoritma prediktif terhadap efisiensi biaya logistik perusahaan.\n"
            )

        extra_p = (
            "Pendekatan analitik modern pada manajemen operasi menggabungkan pencatatan transaksi waktu nyata dengan model peramalan berbasis data historis agar setiap perubahan pola permintaan dapat dideteksi lebih dini oleh pengelola pabrik. "
            "Dalam praktiknya, sistem informasi terpadu menyederhanakan koordinasi antarunit kerja karena data persediaan bahan baku, jadwal produksi harian, dan kapasitas armada pengiriman tersimpan pada satu sumber kebenaran yang sama. "
            "Dengan demikian, perusahaan tidak lagi bergantung pada rekapitulasi manual yang lambat dan rentan terhadap selisih pencatatan ketika menghadapi lonjakan pesanan musiman dari pelanggan.\n\n"
        )
        with open(os.path.join(drafts_dir, "bab_02.md"), "w", encoding="utf-8") as f:
            f.write(
                "## 2.1 Arsitektur Penerapan AI dalam Rantai Pasok Manufaktur\n\n"
                "Integrasi ***artificial intelligence*** pada jaringan logistik memungkinkan sinkronisasi data pengadaan, produksi, dan distribusi secara waktu nyata (Anggraini & Putra, 2023; Mahendra, 2023). "
                "Menurut Pratama & Nugroho (2023), penerapan algoritma *machine learning* mampu menekan distorsi informasi permintaan hingga tiga puluh persen dibandingkan metode konvensional (Budiman, 2024; Gunawan & Lestari, 2023). "
                "Kemampuan analitik prediktif tersebut membantu manajer operasional mengambil keputusan pengadaan secara presisi sekaligus memitigasi risiko gangguan pasokan dan keamanan data (Fauzan, 2024; Hidayat, 2022; Wijaya & Rahmawati, 2024).\n\n"
                + (extra_p * 11)
                + "> Transformasi rantai pasok berbasis algoritma prediktif menuntut integrasi data lintas departemen secara waktu nyata, sehingga setiap distorsi permintaan di tingkat hilir dapat langsung dikompensasi oleh penjadwalan produksi di tingkat hulu tanpa menimbulkan efek cambuk stok yang merugikan (Christopher, 2022: 84).\n\n"
                "![Arsitektur AI](assets/arsitektur_ai.png)\n"
                "**Gambar 2.1** Arsitektur Integrasi AI pada Rantai Pasok\n"
                "*Sumber:* Pratama & Nugroho (2023)\n\n"
                "## 2.2 Dampak Integrasi Algoritma Prediktif terhadap Efisiensi Logistik\n\n"
                "Evaluasi kinerja operasional menunjukkan peningkatan akurasi peramalan yang signifikan setelah adopsi sistem cerdas pada gudang distribusi (Dewi & Hartono, 2021). "
                "Data empiris membuktikan bahwa otomatisasi pemantauan stok menurunkan risiko kelebihan persediaan sekaligus mempercepat waktu tanggap pelanggan dan rute pengiriman (Kusuma & Santoso, 2025; Pratama & Nugroho, 2023). "
                "Tabel 2.1 di bawah ini merangkum perbandingan indikator kinerja antara pendekatan konvensional dan pendekatan berbasis AI.\n\n"
                + (extra_p * 11)
                + "**Tabel 2.1** Perbandingan Kinerja Rantai Pasok Konvensional dan Berbasis AI\n"
                "| No. | Indikator Kinerja | Sistem Konvensional | Sistem Berbasis AI |\n"
                "| :--- | :--- | :--- | :--- |\n"
                "| 1 | Akurasi Peramalan | 68,5% | 91,2% |\n"
                "| 2 | Waktu Respons Stok | 14 hari | 3 hari |\n"
                "*Sumber:* Pratama & Nugroho (2023)\n"
            )

        with open(os.path.join(drafts_dir, "bab_03.md"), "w", encoding="utf-8") as f:
            f.write(
                "## 3.1 Kesimpulan\n\n"
                "1. Arsitektur penerapan *artificial intelligence* dalam rantai pasok industri manufaktur bekerja melalui integrasi data lintas departemen dan pemodelan *machine learning* untuk memprediksi kebutuhan secara akurat.\n"
                "2. Integrasi algoritma prediktif terbukti meningkatkan akurasi peramalan hingga 91,2% dan memangkas waktu respons stok menjadi tiga hari sehingga efisiensi logistik meningkat signifikan.\n\n"
                "## 3.2 Saran\n\n"
                "1. **Bagi Instansi dan Regulator**: Kementerian terkait disarankan menyusun standar interoperabilitas data industri nasional serta memberikan insentif infrastruktur komputasi bagi industri manufaktur menengah.\n"
                "2. **Bagi Masyarakat dan Praktisi Industri**: Manajer operasional perusahaan manufaktur disarankan melakukan audit kesiapan data historis sebelum mengimplementasikan modul peramalan berbasis *machine learning*.\n"
                "3. **Bagi Peneliti Selanjutnya**: Penelitian mendatang disarankan memperluas cakupan analisis pada aspek keamanan siber rantai pasok serta perhitungan *return on investment* jangka panjang.\n"
            )

        res = compile_workspace(tmp_dir, margin_override="4-4-3-3")
        docx_p = res["docx_path"]
        pdf_p = res["pdf_path"]

        assert cfg["page"]["margin_top_cm"] == 4.0, "CLI margin_override gagal mengalahkan blueprint"
        assert os.path.exists(docx_p) and os.path.getsize(docx_p) > 5000, "File DOCX gagal dibuat"
        assert os.path.exists(pdf_p) and os.path.getsize(pdf_p) > 2000, "File PDF gagal dibuat"
        assert len(res["toc"]["lot"]) == 1, "Daftar Tabel TOC gagal tercatat"
        assert len(res["toc"]["lof"]) == 1, "Daftar Gambar TOC gagal tercatat"
        assert len(res["audit"]["warnings"]) == 0, f"Ada peringatan tak terduga: {res['audit']['warnings']}"
        # Verifikasi aturan jumlah referensi wajib > 10 (10 referensi harus memicu warning)
        too_few_audit = audit_makalah({**sample_bp, "references": sample_bp["references"][:10]}, {}, tmp_dir)
        assert any("[Jumlah Referensi]" in w for w in too_few_audit["warnings"]), "Linter gagal mendeteksi jumlah referensi <= 10!"
        thin_bab2 = parse_markdown_blocks("## 2.1 Pendahuluan Topik\n\nSatu paragraf pendek saja. Kalimat dua. Kalimat tiga.\n")
        thin_audit = audit_makalah(sample_bp, {"BAB II": thin_bab2}, tmp_dir)
        assert any("[Kedalaman Pembahasan]" in w for w in thin_audit["warnings"]), "Linter gagal mendeteksi Bab II tipis!"
        assert any("[Kedalaman Sub-bab]" in w for w in thin_audit["warnings"]), "Linter gagal mendeteksi sub-bab < 3 paragraf!"
        toc_entries = res["toc"]["toc"]
        assert toc_entries[0]["page"] == "i", f"Halaman awal pertama seharusnya 'i', bukan '{toc_entries[0]['page']}'"
        bab1_entries = [e for e in toc_entries if e["label"].startswith("BAB I ")]
        assert bab1_entries and bab1_entries[0]["page"] == "1", f"BAB I seharusnya mulai halaman '1', bukan '{bab1_entries}'"
        assert all(e["page"] != "0" for e in toc_entries), f"Ditemukan nomor halaman '0' di TOC: {toc_entries}"
        # Verifikasi deduplikasi kota dan deteksi placeholder sampul
        dedup_meta = get_cover_metadata({**sample_bp, "institution": "Institut Teknologi Bandung", "city": "Bandung"})
        assert dedup_meta["inst_lines"].count("BANDUNG") == 0 and dedup_meta["inst_lines"][-2] == "INSTITUT TEKNOLOGI BANDUNG"
        ph_audit = audit_makalah({**sample_bp, "lecturer": "[Nama Dosen Pengampu]"}, {}, tmp_dir)
        assert any("[Komponen Sampul]" in w for w in ph_audit["warnings"]), "Linter gagal mendeteksi placeholder sampul!"
        cov_spec_check = _build_cover_zone_spec(sample_bp, res["logo_path"])
        assert sum(cov_spec_check["row_heights"]) <= cov_spec_check["cov_frame_h"] + 1.0, "Tinggi Grid 3-Zona sampul melampaui frame!"

        with zipfile.ZipFile(docx_p, "r") as zf:
            xml_bytes = zf.read("word/document.xml")
            settings_bytes = zf.read("word/settings.xml")
            assert b"w:sectPr" in xml_bytes, "Section break XML tidak ditemukan di DOCX"
            assert b"w:pgNumType" in xml_bytes, "Penomoran halaman XML tidak ditemukan di DOCX"
            assert b"w:tblHeader" in xml_bytes, "Header tabel berulang (w:tblHeader) tidak ditemukan"
            assert b"w:cantSplit" in xml_bytes, "Proteksi baris tabel (w:cantSplit) tidak ditemukan"
            assert b'w:vAlign w:val="bottom"' in xml_bytes, "Penjangkaran bawah blok institusi cover (w:vAlign bottom) tidak ditemukan"
            assert b'w:tblLayout w:type="fixed"' in xml_bytes, "Fixed table layout Word tidak ditemukan"
            assert b"w:tblCellMar" in xml_bytes, "Padding sel tabel Word (w:tblCellMar) tidak ditemukan"
            assert b'w:insideH w:val="single"' in xml_bytes and b'w:insideV w:val="single"' in xml_bytes, "Garis kotak penuh tabel Word tidak ditemukan"
            assert b'w:fill="F2F2F2"' in xml_bytes, "Shading header tabel Word tidak ditemukan"
            assert b"w:sdt" in xml_bytes and b'TOC \\o "1-3"' in xml_bytes, "Native Word TOC field tidak ditemukan di DOCX"
            assert b"Caption Tabel" in xml_bytes and b"Caption Gambar" in xml_bytes, "Style Caption Tabel/Gambar TOC tidak ditemukan"
            assert b"w:updateFields" in settings_bytes, "Auto-update fields di word/settings.xml tidak aktif"
            # Pastikan paragraf prosa "Tabel 2.1 di bawah ini..." TIDAK terhapus!
            assert b"Tabel 2.1 di bawah ini merangkum" in xml_bytes, "Paragraf berawalan 'Tabel 2.1' terhapus oleh parser!"

        # Verifikasi aspek rasio logo non-persegi tetap terjaga
        rect_logo = os.path.join(assets_dir, "logo_kampus.png")
        Image.new("RGBA", (400, 200), (20, 42, 78, 255)).save(rect_logo, "PNG")
        rw, rh = compute_cover_dimensions_cm(rect_logo, 4.5)
        assert abs(rw - 4.5) < 0.01 and abs(rh - 2.25) < 0.01, "Rasio logo kampus terdistorsi!"
        assert resolve_cover_logo({"logo_path": ""}, tmp_dir, res["logo_path"]) == rect_logo, "Auto-deteksi logo_kampus.png gagal!"

        with open(pdf_p, "rb") as pf:
            header = pf.read(5)
            assert header == b"%PDF-", "Header PDF tidak valid"

        print(f"[OK] DOCX Valid : {docx_p} ({os.path.getsize(docx_p)} bytes)")
        print(f"[OK] PDF Valid  : {pdf_p} ({os.path.getsize(pdf_p)} bytes)")
        print(f"[OK] Margin CLI : {cfg['page']['margin_preset']}")
        print(f"[OK] TOC Items  : {len(res['toc']['toc'])} entri | Tabel: {len(res['toc']['lot'])} | Gambar: {len(res['toc']['lof'])}")
        print(f"[OK] Linter Warn: {len(res['audit']['warnings'])} peringatan")
        print("=== [SMOKE TEST PASSED] ===")
        return 0
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Kompilator Makalah Akademik Standar (DOCX & PDF)")
    parser.add_argument("--workspace", help="Path ke folder makalah_workspace/<slug>")
    parser.add_argument("--output-dir", help="Path folder output untuk menyimpan DOCX dan PDF")
    parser.add_argument("--margin", help="Override batas margin Kiri,Atas,Kanan,Bawah (contoh: 4,3,3,3 atau 4,4,3,3)")
    parser.add_argument("--smoke-test", action="store_true", help="Jalankan pengujian mandiri (smoke test)")
    args = parser.parse_args()

    if args.margin:
        apply_margin_preset(args.margin)

    if args.smoke_test:
        return run_smoke_test()

    if not args.workspace:
        parser.error("Argumen --workspace wajib diisi (atau gunakan --smoke-test).")

    result = compile_workspace(args.workspace, args.output_dir, margin_override=args.margin)
    print(f"[SUCCESS] DOCX berhasil dibuat: {result['docx_path']}")
    print(f"[SUCCESS] PDF berhasil dibuat : {result['pdf_path']}")
    audit = result["audit"]
    print(
        f"[STATS] Paragraf: {audit['stats']['total_paragraphs']} | "
        f"Tabel: {audit['stats']['total_tables']} | "
        f"Gambar: {audit['stats']['total_figures']} | "
        f"Referensi: {audit['stats']['total_references']}"
    )
    if audit["warnings"]:
        print("[ACADEMIC LINTER WARNINGS]:")
        for w in audit["warnings"]:
            print(f"  - {w}")
    else:
        print("[ACADEMIC LINTER]: 100% Patuh Standar Makalah Akademik (0 Peringatan).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
