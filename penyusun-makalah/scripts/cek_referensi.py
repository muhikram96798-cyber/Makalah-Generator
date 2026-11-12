#!/usr/bin/env python3
"""
Verifikator Keterbukaan Referensi Ilmiah & Anti-Paywall (penyusun-makalah)
Menguji apakah URL/DOI jurnal benar-benar terbuka (OPEN_ACCESS / PDF dapat diakses)
atau terkunci paywall berbayar (PAYWALLED) maupun gagal diakses (UNREACHABLE).
"""

cfg = {
    "min_open_references": 11,  # Aturan wajib di atas 10 referensi terbuka
    "max_extract_chars": 4500,
    "max_pdf_bytes": 15728640,  # 15 MB
    "network": {
        "timeout_seconds": 15,
        "max_html_bytes": 524288,
        "user_agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
    },
    "paywall_indicators": [
        "buy this article",
        "purchase pdf",
        "purchase this article",
        "get access",
        "institutional access",
        "sign in via your institution",
        "access through your institution",
        "rent or buy",
        "subscribe to journal",
        "subscribe to read",
        "pay-per-view",
        "locked content",
        "check access",
    ],
    "open_access_indicators": [
        "open access",
        "creative commons",
        "cc-by",
        "full text pdf",
        "download pdf",
        "unduh pdf",
        "citation_pdf_url",
    ],
}

import argparse
import http.cookiejar
import json
import os
import re
import ssl
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request

try:
    import pymupdf
except ImportError:
    pymupdf = None


def normalize_reference_url(url_or_doi: str) -> str:
    raw = (url_or_doi or "").strip()
    if not raw:
        return ""
    if raw.lower().startswith("doi:"):
        raw = raw[4:].strip()
    if re.match(r"^10\.\d{4,9}/\S+$", raw):
        return f"https://doi.org/{raw}"
    if not re.match(r"^https?://", raw, re.IGNORECASE):
        return f"https://{raw}"
    return raw


def _build_opener() -> urllib.request.OpenerDirector:
    cj = http.cookiejar.CookieJar()
    ctx = ssl.create_default_context()
    return urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cj),
        urllib.request.HTTPSHandler(context=ctx),
    )


def _fetch_url(url: str, opener: urllib.request.OpenerDirector, method: str = "GET") -> tuple[int, str, str, bytes]:
    req = urllib.request.Request(
        url,
        method=method,
        headers={
            "User-Agent": cfg["network"]["user_agent"],
            "Accept": "text/html,application/pdf,application/xhtml+xml,*/*;q=0.8",
            "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
        },
    )
    timeout = cfg["network"]["timeout_seconds"]
    with opener.open(req, timeout=timeout) as resp:
        code = resp.getcode() or 200
        final_url = resp.geturl()
        ctype = (resp.headers.get("Content-Type") or "").lower()
        body = b""
        if method != "HEAD":
            body = resp.read(cfg["network"]["max_html_bytes"])
        return code, final_url, ctype, body


def _extract_pdf_candidates(html_text: str, base_url: str) -> list[str]:
    candidates: list[str] = []
    # 1. Meta tag standar Google Scholar / OJS / Digital Commons (urutan atribut bebas)
    meta_patterns = (
        r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']citation_pdf_url["\']',
    )
    for pat in meta_patterns:
        for m in re.finditer(pat, html_text, re.IGNORECASE):
            candidates.append(urllib.parse.urljoin(base_url, m.group(1).strip()))
    # 2. Tautan PDF atau endpoint OJS / Digital Commons (/cgi/viewcontent.cgi)
    for m in re.finditer(r'href=["\']([^"\']+)["\']', html_text, re.IGNORECASE):
        href = m.group(1).strip()
        if re.search(r"\.pdf($|\?)|/article/download/|/download/|/cgi/viewcontent\.cgi|pdf", href, re.IGNORECASE):
            if not href.startswith("#") and not href.lower().startswith("javascript:"):
                candidates.append(urllib.parse.urljoin(base_url, href))
    # Deduplikasi sambil menjaga urutan
    seen = set()
    unique = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique[:5]


def verify_single_reference(url_or_doi: str, citation: str = "", fetcher=None) -> dict:
    """Uji satu URL/DOI referensi dan klasifikasikan sebagai OPEN_ACCESS, PAYWALLED, atau UNREACHABLE."""
    fetch_fn = fetcher or _fetch_url
    norm_url = normalize_reference_url(url_or_doi)
    if not norm_url:
        return {
            "citation": citation,
            "url": url_or_doi,
            "final_url": "",
            "status": "UNREACHABLE",
            "access_type": "EMPTY_URL",
            "pdf_url": "",
            "http_code": 0,
            "notes": "URL atau DOI kosong.",
        }

    opener = _build_opener()
    try:
        code, final_url, ctype, raw_bytes = fetch_fn(norm_url, opener, method="GET")
    except urllib.error.HTTPError as he:
        status_label = "PAYWALLED" if he.code in (401, 402, 403) else "UNREACHABLE"
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": he.geturl() or norm_url,
            "status": status_label,
            "access_type": f"HTTP_{he.code}",
            "pdf_url": "",
            "http_code": he.code,
            "notes": f"Server menolak akses atau file tidak ditemukan (HTTP {he.code}).",
        }
    except Exception as exc:
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": norm_url,
            "status": "UNREACHABLE",
            "access_type": "NETWORK_ERROR",
            "pdf_url": "",
            "http_code": 0,
            "notes": f"Gagal mengakses URL: {exc}",
        }

    # Kasus 1: Langsung mengembalikan file PDF valid
    if "application/pdf" in ctype or raw_bytes.startswith(b"%PDF-"):
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": final_url,
            "status": "OPEN_ACCESS",
            "access_type": "DIRECT_PDF",
            "pdf_url": final_url,
            "http_code": code,
            "notes": "PDF langsung terbuka dan valid.",
        }

    # Kasus 2: Halaman HTML artikel -> cek PDF link terbuka & indikator paywall
    html_text = raw_bytes.decode("utf-8", errors="replace")
    lower_html = html_text.lower()

    pdf_candidates = _extract_pdf_candidates(html_text, final_url)
    verified_pdf = ""
    for cand in pdf_candidates:
        try:
            p_code, p_final, p_ctype, p_bytes = fetch_fn(cand, opener, method="GET")
            if p_code == 200 and ("application/pdf" in p_ctype or p_bytes.startswith(b"%PDF-")):
                verified_pdf = p_final
                break
        except Exception:
            continue

    if verified_pdf:
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": final_url,
            "status": "OPEN_ACCESS",
            "access_type": "HTML_WITH_OPEN_PDF",
            "pdf_url": verified_pdf,
            "http_code": code,
            "notes": "Halaman artikel terbuka dengan PDF full-text terverifikasi.",
        }

    paywall_hits = [kw for kw in cfg["paywall_indicators"] if kw in lower_html]
    if paywall_hits:
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": final_url,
            "status": "PAYWALLED",
            "access_type": "PAYWALL_DETECTED",
            "pdf_url": "",
            "http_code": code,
            "notes": f"Terindikasi paywall/berbayar ({', '.join(paywall_hits[:3])}).",
        }

    open_hits = [kw for kw in cfg["open_access_indicators"] if kw in lower_html]
    text_only = re.sub(r"<[^>]+>", " ", html_text)
    word_count = len(text_only.split())
    if open_hits or word_count >= 450:
        return {
            "citation": citation,
            "url": norm_url,
            "final_url": final_url,
            "status": "OPEN_ACCESS",
            "access_type": "OPEN_HTML_FULLTEXT",
            "pdf_url": "",
            "http_code": code,
            "notes": f"Halaman artikel HTML terbuka (panjang ~{word_count} kata).",
        }

    return {
        "citation": citation,
        "url": norm_url,
        "final_url": final_url,
        "status": "PAYWALLED",
        "access_type": "RESTRICTED_OR_ABSTRACT_ONLY",
        "pdf_url": "",
        "http_code": code,
        "notes": "Hanya memuat ringkasan pendek tanpa tautan PDF terbuka.",
    }


def verify_batch(items: list[dict], output_path: str | None = None, fetcher=None) -> list[dict]:
    results = []
    for item in items:
        cit = item.get("citation", "")
        url = item.get("url") or item.get("doi") or ""
        res = verify_single_reference(url, cit, fetcher=fetcher)
        results.append(res)
    if output_path:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
    return results


def _fetch_binary(url: str, opener: urllib.request.OpenerDirector, max_bytes: int) -> bytes:
    req = urllib.request.Request(
        url,
        method="GET",
        headers={
            "User-Agent": cfg["network"]["user_agent"],
            "Accept": "application/pdf,text/html,*/*;q=0.8",
        },
    )
    with opener.open(req, timeout=cfg["network"]["timeout_seconds"] + 10) as resp:
        return resp.read(max_bytes)


def _clean_extracted_text(raw_text: str) -> str:
    text = re.sub(r"[ \t]+", " ", raw_text.replace("\r\n", "\n"))
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _slice_key_sections(full_text: str, max_chars: int) -> dict:
    """Ambil inti isi jurnal (pembuka/abstrak, metode & hasil, kesimpulan) agar ringkas dan padat fakta."""
    cleaned = _clean_extracted_text(full_text)
    total_len = len(cleaned)
    if total_len <= max_chars:
        return {"excerpt": cleaned, "total_chars": total_len}

    head = cleaned[: int(max_chars * 0.40)]
    mid_start = max(0, (total_len // 2) - int(max_chars * 0.175))
    mid = cleaned[mid_start : mid_start + int(max_chars * 0.35)]
    tail = cleaned[-int(max_chars * 0.25) :]
    combined = f"[BAGIAN AWAL / ABSTRAK]\n{head}\n\n[BAGIAN TENGAH / METODE & HASIL]\n{mid}\n\n[BAGIAN AKHIR / KESIMPULAN]\n{tail}"
    return {"excerpt": combined[: max_chars + 120], "total_chars": total_len}


def extract_journal_notes(verified_items: list[dict], output_path: str, binary_fetcher=None) -> list[dict]:
    """Ekstrak inti isi PDF/HTML dari jurnal berstatus OPEN_ACCESS ke references_notes.json."""
    opener = _build_opener()
    fetch_bin = binary_fetcher or (lambda u: _fetch_binary(u, opener, cfg.get("max_pdf_bytes", 15728640)))
    max_chars = cfg.get("max_extract_chars", 4500)
    notes: list[dict] = []

    for item in verified_items:
        if item.get("status") != "OPEN_ACCESS":
            continue
        cit = item.get("citation", "")
        target_url = item.get("pdf_url") or item.get("final_url") or item.get("url") or ""
        entry = {
            "citation": cit,
            "source_url": target_url,
            "access_type": item.get("access_type", ""),
            "page_count": 0,
            "total_chars": 0,
            "excerpt": "",
            "extract_status": "FAILED",
        }
        try:
            raw_bytes = fetch_bin(target_url)
            if raw_bytes.startswith(b"%PDF-") and pymupdf is not None:
                doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
                pages_txt = [page.get_text("text") for page in doc]
                entry["page_count"] = len(doc)
                doc.close()
                sliced = _slice_key_sections("\n".join(pages_txt), max_chars)
                entry["excerpt"] = sliced["excerpt"]
                entry["total_chars"] = sliced["total_chars"]
                entry["extract_status"] = "OK_PDF" if entry["excerpt"] else "EMPTY_PDF_TEXT"
            else:
                html_str = raw_bytes.decode("utf-8", errors="replace")
                html_str = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html_str, flags=re.DOTALL | re.IGNORECASE)
                plain = re.sub(r"<[^>]+>", " ", html_str)
                sliced = _slice_key_sections(plain, max_chars)
                entry["excerpt"] = sliced["excerpt"]
                entry["total_chars"] = sliced["total_chars"]
                entry["extract_status"] = "OK_HTML" if entry["excerpt"] else "EMPTY_HTML_TEXT"
        except Exception as exc:
            entry["excerpt"] = f"Gagal mengekstrak isi: {exc}"
        notes.append(entry)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(notes, f, indent=2, ensure_ascii=False)
    return notes


def run_smoke_test() -> int:
    """Uji logika anti-paywall & ekstraksi teks PDF pymupdf secara luring (sandbox-safe)."""
    print("=== RUNNING MANDATORY SMOKE TEST (cek_referensi) ===")

    sample_pdf_bytes = b"%PDF-1.4\n%Mock Open Access PDF\n%%EOF"
    if pymupdf is not None:
        pdf_doc = pymupdf.open()
        pg = pdf_doc.new_page()
        pg.insert_text((72, 72), "Abstrak: Chatbot berbasis NLP dengan Naive Bayes mencapai akurasi 92 persen pada layanan akademik.")
        sample_pdf_bytes = pdf_doc.tobytes()
        pdf_doc.close()

    def mock_fetcher(url: str, opener, method: str = "GET"):
        if url.endswith("/open.pdf"):
            return 200, url, "application/pdf", sample_pdf_bytes
        if url.endswith("/ojs-article"):
            html_body = (
                '<html><head><meta name="citation_pdf_url" content="https://journal.example.ac.id/open.pdf"></head>'
                "<body><h1>Artikel Terbuka OJS</h1><a href='https://journal.example.ac.id/open.pdf'>Download PDF</a></body></html>"
            ).encode("utf-8")
            return 200, url, "text/html; charset=utf-8", html_body
        if url.endswith("/paywall"):
            html_body = (
                "<html><body><h1>Article Preview</h1>"
                "<p>Buy this article or Sign in via your institution to get access ($39.95).</p></body></html>"
            ).encode("utf-8")
            return 200, url, "text/html; charset=utf-8", html_body
        if url.endswith("/forbidden"):
            raise urllib.error.HTTPError(url, 403, "Forbidden", {}, None)
        raise urllib.error.HTTPError(url, 404, "Not Found", {}, None)

    tmp_out = os.path.join(tempfile.gettempdir(), "smoke_refs_verified.json")
    tmp_notes = os.path.join(tempfile.gettempdir(), "smoke_refs_notes.json")
    try:
        items = [
            {"citation": "Penulis A (2024)", "url": "https://journal.example.ac.id/open.pdf"},
            {"citation": "Penulis B (2025)", "url": "https://journal.example.ac.id/ojs-article"},
            {"citation": "Penulis C (2023)", "url": "https://publisher.example.com/paywall"},
            {"citation": "Penulis D (2022)", "url": "https://publisher.example.com/forbidden"},
            {"citation": "Penulis E (2024)", "doi": "10.23969/infomatek.v26i2.19007"},
        ]
        # Uji normalisasi DOI & Digital Commons meta terbalik sebelum batch
        assert normalize_reference_url(items[4]["doi"]) == "https://doi.org/10.23969/infomatek.v26i2.19007"
        dc_meta = '<meta content="https://scholarhub.ui.ac.id/cgi/viewcontent.cgi?article=6&context=jkskn" name="citation_pdf_url">'
        assert _extract_pdf_candidates(dc_meta, "https://scholarhub.ui.ac.id/")[0].endswith("context=jkskn")
        results = verify_batch(items[:4], output_path=tmp_out, fetcher=mock_fetcher)
        for r in results:
            print(f"  [{r['status']}] {r['citation']} -> {r['access_type']} ({r['notes']})")
        assert results[0]["status"] == "OPEN_ACCESS" and results[0]["access_type"] == "DIRECT_PDF"
        assert results[1]["status"] == "OPEN_ACCESS" and results[1]["access_type"] == "HTML_WITH_OPEN_PDF"
        assert results[2]["status"] == "PAYWALLED" and results[2]["access_type"] == "PAYWALL_DETECTED"
        assert results[3]["status"] == "PAYWALLED" and results[3]["http_code"] == 403
        assert os.path.exists(tmp_out)

        if pymupdf is not None:
            notes = extract_journal_notes(results[:2], tmp_notes, binary_fetcher=lambda u: sample_pdf_bytes)
            assert len(notes) == 2 and notes[0]["extract_status"] == "OK_PDF"
            assert "92 persen" in notes[0]["excerpt"], "Ekstraksi teks PDF gagal membaca isi!"
            print(f"  [EXTRACT OK] {len(notes)} catatan PDF diekstrak via pymupdf.")

        print("=== [SMOKE TEST PASSED] ===")
        return 0
    finally:
        for p in (tmp_out, tmp_notes):
            if os.path.exists(p):
                os.remove(p)


def main() -> int:
    parser = argparse.ArgumentParser(description="Verifikator Keterbukaan Jurnal, Anti-Paywall, & Ekstraktor Isi PDF")
    parser.add_argument("--input", help="File JSON daftar kandidat referensi ([{citation, url/doi}])")
    parser.add_argument("--from-verified", help="Langsung baca references_verified.json yang sudah ada untuk ekstraksi isi")
    parser.add_argument("--output", help="Path output JSON hasil verifikasi (mis. references_verified.json)")
    parser.add_argument("--extract-notes", help="Path output JSON inti isi jurnal (mis. references_notes.json)")
    parser.add_argument("--urls", nargs="+", help="Daftar URL/DOI yang ingin diuji langsung")
    parser.add_argument("--allow-fewer", action="store_true", help="Izinkan jumlah referensi <= 10 (khusus uji parsial)")
    parser.add_argument("--smoke-test", action="store_true", help="Jalankan pengujian mandiri luring")
    args = parser.parse_args()

    if args.smoke_test:
        return run_smoke_test()

    if args.from_verified:
        with open(args.from_verified, "r", encoding="utf-8") as vf:
            results = json.load(vf)
    else:
        items: list[dict] = []
        if args.input:
            with open(args.input, "r", encoding="utf-8") as f:
                items = json.load(f)
        elif args.urls:
            items = [{"citation": f"Referensi {idx}", "url": u} for idx, u in enumerate(args.urls, 1)]
        else:
            parser.error("Gunakan --input <file.json>, --from-verified <verified.json>, --urls <url...>, atau --smoke-test.")
        results = verify_batch(items, output_path=args.output)

    open_count = sum(1 for r in results if r.get("status") == "OPEN_ACCESS")
    blocked_count = len(results) - open_count

    if not args.from_verified:
        for r in results:
            print(f"[{r['status']}] {r['citation'] or r['url']} | {r['access_type']} | {r['notes']}")
    print(f"[SUMMARY] Terbuka: {open_count} | Ditolak (Paywall/Gagal): {blocked_count}")

    if args.extract_notes:
        notes = extract_journal_notes(results, args.extract_notes)
        ok_notes = sum(1 for n in notes if n["extract_status"].startswith("OK_"))
        print(f"[EXTRACT] Berhasil mengekstrak isi {ok_notes}/{len(notes)} jurnal ke: {args.extract_notes}")

    min_req = cfg.get("min_open_references", 11)
    if (args.input or args.from_verified) and not args.allow_fewer and open_count < min_req:
        print(f"[WARNING] Jumlah referensi OPEN_ACCESS ({open_count}) harus di atas 10 (minimal {min_req} sumber).")
        return 3 if blocked_count == 0 else 2
    return 0 if blocked_count == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
