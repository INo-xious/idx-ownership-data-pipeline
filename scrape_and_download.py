"""Scrape IDX Keterbukaan Informasi for KSEI 5% ownership attachments.

The IDX page blocks simple HTTP clients, so runtime collection uses a real
Playwright browser session. The pure parsing helpers stay dependency-light so
tests can validate filename/date/manifest behavior without browser packages.
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import json
import os
import re
import subprocess
import sys
import time
import urllib.parse
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import requests
from tqdm import tqdm


IDX_URL = "https://www.idx.co.id/id/perusahaan-tercatat/keterbukaan-informasi"
DEFAULT_KEYWORD = "Pemegang Saham di atas 5% (KSEI)"

PDF_RE = re.compile(r"\.pdf(?:\?|$)", re.IGNORECASE)
TARGET_ATTACHMENT_RE = re.compile(r"^20\d{6}_.+_lamp\d+\.pdf$", re.IGNORECASE)
FILENAME_DATE_RE = re.compile(r"^(20\d{2})(\d{2})(\d{2})_")
DISCLOSURE_DATETIME_RE = re.compile(
    r"(\d{1,2})\s+"
    r"(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)"
    r"\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})",
    re.IGNORECASE,
)

INDONESIAN_MONTHS = {
    "januari": 1,
    "februari": 2,
    "maret": 3,
    "april": 4,
    "mei": 5,
    "juni": 6,
    "juli": 7,
    "agustus": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "desember": 12,
}


@dataclass
class DisclosureAttachment:
    disclosure_title: str
    disclosure_datetime: str
    announcement_date: str
    detail_url: str
    attachment_filename: str
    attachment_url: str
    filename_date: str
    local_pdf_path: str = ""
    download_status: str = "pending"
    error: str = ""


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def parse_indonesian_datetime(value: str) -> str:
    """Return ISO-like local datetime for IDX strings such as 08 Juni 2026."""
    m = DISCLOSURE_DATETIME_RE.search(_norm(value))
    if not m:
        return ""
    day, month_name, year, hh, mm, ss = m.groups()
    month = INDONESIAN_MONTHS.get(month_name.lower())
    if not month:
        return ""
    return (
        f"{int(year):04d}-{month:02d}-{int(day):02d} "
        f"{int(hh):02d}:{int(mm):02d}:{int(ss):02d}"
    )


def parse_iso_date(value: str) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(value[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def filename_date(filename: str) -> str:
    m = FILENAME_DATE_RE.search(os.path.basename(filename or ""))
    if not m:
        return ""
    y, mth, d = m.groups()
    try:
        return date(int(y), int(mth), int(d)).isoformat()
    except ValueError:
        return ""


def is_target_attachment_filename(filename: str) -> bool:
    return bool(TARGET_ATTACHMENT_RE.match(os.path.basename(filename or "").strip()))


def safe_filename(url: str, preferred_name: str | None = None) -> str:
    name = _norm(preferred_name or "")
    if not name:
        path = urllib.parse.urlparse(url).path
        name = os.path.basename(path) or "document.pdf"
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return re.sub(r"[^\w\-. ()]", "_", name)[:180]


def looks_like_pdf(data: bytes) -> bool:
    return bool(data and len(data) >= 1024 and data[:4] == b"%PDF")


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Scrape IDX KI KSEI 5% ownership attachments and optionally extract them"
    )
    p.add_argument(
        "--keyword",
        default=DEFAULT_KEYWORD,
        help='Search keyword typed into IDX "Kata kunci".',
    )
    p.add_argument(
        "--out-pdf-dir",
        default="outputs/pdfs",
        help="Folder to save downloaded PDFs.",
    )
    p.add_argument(
        "--manifest-out",
        default="",
        help="Manifest JSON path. Defaults to <out-pdf-dir>/download_manifest.json.",
    )
    p.add_argument("--date-from", default="", help="Optional YYYY-MM-DD lower bound.")
    p.add_argument("--date-to", default="", help="Optional YYYY-MM-DD upper bound.")
    p.add_argument(
        "--max-pdfs",
        type=int,
        default=0,
        help="If >0, stop after this many target PDFs.",
    )
    p.add_argument(
        "--extract",
        action="store_true",
        help="After download, run extract_ownership_table.py on each PDF.",
    )
    p.add_argument(
        "--extract-out-dir",
        default="outputs/extracted",
        help="Where to save extracted XLSX files.",
    )
    p.add_argument(
        "--extract-debug-dir",
        default="",
        help="If set, write debug artifacts per PDF under this folder.",
    )
    p.add_argument(
        "--overwrite",
        action="store_true",
        help="Re-download / re-extract even if output files already exist.",
    )
    return p


def _container_excerpt(text: str, anchor_text: str) -> str:
    text = _norm(text)
    if len(text) <= 1200:
        return text
    idx = text.lower().find(anchor_text.lower())
    if idx < 0:
        return text[:1200]
    start = max(0, idx - 600)
    end = min(len(text), idx + len(anchor_text) + 600)
    return text[start:end]


def _title_from_container(text: str) -> str:
    lines = [_norm(x) for x in re.split(r"[\r\n]+", text) if _norm(x)]
    for line in lines:
        if DEFAULT_KEYWORD.lower() in line.lower():
            return line
    compact = _norm(text)
    idx = compact.lower().find(DEFAULT_KEYWORD.lower())
    if idx >= 0:
        end = compact.find("@", idx)
        if end < 0:
            end = min(len(compact), idx + 180)
        return _norm(compact[idx:end])
    return ""


def normalize_disclosure_candidate(raw: dict[str, Any]) -> DisclosureAttachment | None:
    href = str(raw.get("href") or "")
    text = _norm(str(raw.get("text") or ""))
    container_text = _container_excerpt(str(raw.get("container_text") or ""), text)

    filename = safe_filename(href, text)
    if not PDF_RE.search(href) and not filename.lower().endswith(".pdf"):
        return None
    if not is_target_attachment_filename(filename):
        return None

    disclosure_datetime = parse_indonesian_datetime(container_text)
    announcement_date = disclosure_datetime[:10] if disclosure_datetime else ""
    fdate = filename_date(filename)
    title = _title_from_container(container_text)

    detail_url = str(raw.get("detail_url") or "")
    if not detail_url:
        detail_url = str(raw.get("page_url") or "")

    return DisclosureAttachment(
        disclosure_title=title,
        disclosure_datetime=disclosure_datetime,
        announcement_date=announcement_date or fdate,
        detail_url=detail_url,
        attachment_filename=filename,
        attachment_url=href,
        filename_date=fdate,
    )


def dedupe_attachments(items: Iterable[DisclosureAttachment]) -> list[DisclosureAttachment]:
    """Deduplicate target attachments, preferring richer metadata."""
    by_url: dict[str, DisclosureAttachment] = {}
    for item in items:
        key = item.attachment_url or item.attachment_filename
        existing = by_url.get(key)
        if existing is None:
            by_url[key] = item
            continue
        old_score = sum(bool(getattr(existing, f)) for f in ("disclosure_title", "disclosure_datetime", "detail_url"))
        new_score = sum(bool(getattr(item, f)) for f in ("disclosure_title", "disclosure_datetime", "detail_url"))
        if new_score > old_score:
            by_url[key] = item
    return sorted(by_url.values(), key=lambda x: (x.announcement_date, x.attachment_filename))


def filter_by_date_range(
    items: Iterable[DisclosureAttachment],
    date_from: str = "",
    date_to: str = "",
) -> list[DisclosureAttachment]:
    start = parse_iso_date(date_from)
    end = parse_iso_date(date_to)
    out: list[DisclosureAttachment] = []
    for item in items:
        item_date = parse_iso_date(item.announcement_date) or parse_iso_date(item.filename_date)
        if start and item_date and item_date < start:
            continue
        if end and item_date and item_date > end:
            continue
        out.append(item)
    return out


def write_manifest(items: list[DisclosureAttachment], manifest_path: Path) -> None:
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    data = [asdict(x) for x in items]
    manifest_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    csv_path = manifest_path.with_suffix(".csv")
    if data:
        with csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(data[0].keys()))
            writer.writeheader()
            writer.writerows(data)


def load_manifest(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def run_extractor(pdf_path: Path, out_xlsx: Path, manifest_path: Path | None, debug_root: str = "") -> None:
    cmd = [
        sys.executable,
        str(Path(__file__).with_name("extract_ownership_table.py")),
        "--pdf",
        str(pdf_path),
        "--out",
        str(out_xlsx),
    ]
    if manifest_path:
        cmd += ["--manifest", str(manifest_path)]
    if debug_root:
        cmd += ["--debug-dir", str(Path(debug_root) / pdf_path.stem)]
    subprocess.run(cmd, check=True)


async def collect_target_attachments(page: Any) -> list[DisclosureAttachment]:
    raw_items = await page.evaluate(
        """() => {
            const normalize = (value) => (value || "").replace(/\\s+/g, " ").trim();
            const anchors = Array.from(document.querySelectorAll("a[href]"));
            return anchors.map((anchor) => {
                const href = anchor.href || anchor.getAttribute("href") || "";
                const text = normalize(anchor.textContent || "");
                let container = anchor;
                for (let i = 0; i < 9 && container && container.parentElement; i++) {
                    const parent = container.parentElement;
                    const parentText = normalize(parent.innerText || parent.textContent || "");
                    if (
                        parentText.includes("Pemegang Saham di atas 5%") ||
                        parentText.includes(text)
                    ) {
                        container = parent;
                    }
                }
                const containerText = normalize(container ? (container.innerText || container.textContent || "") : "");
                const links = container ? Array.from(container.querySelectorAll("a[href]")) : [];
                const detail = links.find((a) => {
                    const t = normalize(a.textContent || "");
                    const h = a.href || a.getAttribute("href") || "";
                    return t.includes("Pemegang Saham di atas 5%") && !/\\.pdf(\\?|$)/i.test(h);
                });
                return {
                    href,
                    text,
                    container_text: containerText,
                    detail_url: detail ? (detail.href || detail.getAttribute("href") || "") : "",
                    page_url: window.location.href
                };
            });
        }"""
    )
    items = []
    for raw in raw_items:
        item = normalize_disclosure_candidate(raw)
        if item:
            items.append(item)
    return items


async def main(args: argparse.Namespace) -> None:
    try:
        from playwright.async_api import async_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright is required for IDX scraping. Install requirements and run "
            "`python -m playwright install chromium`."
        ) from exc

    out_dir = Path(args.out_pdf_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = Path(args.manifest_out) if args.manifest_out else out_dir / "download_manifest.json"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()
        await page.goto(IDX_URL, wait_until="networkidle")

        search_box = await page.query_selector('input[placeholder*="Kata kunci" i]')
        if not search_box:
            raise RuntimeError('Input "Kata kunci" not found on IDX page.')

        await search_box.click()
        await search_box.fill(args.keyword)
        await search_box.press("Enter")
        await page.wait_for_timeout(2500)

        found: list[DisclosureAttachment] = []
        while True:
            found.extend(await collect_target_attachments(page))

            next_selectors = [
                'a[aria-label="Next"]',
                'button[aria-label="Next"]',
                'a:has-text("Berikutnya")',
                'button:has-text("Berikutnya")',
                'a:has-text("Next")',
                'button:has-text("Next")',
                "li.pagination-next a",
            ]
            next_btn = None
            for sel in next_selectors:
                el = await page.query_selector(sel)
                if el:
                    next_btn = el
                    break
            if not next_btn:
                break

            aria_disabled = (await next_btn.get_attribute("aria-disabled")) or ""
            disabled_attr = await next_btn.get_attribute("disabled")
            cls = (await next_btn.get_attribute("class")) or ""
            if disabled_attr is not None or aria_disabled.lower() == "true" or "disabled" in cls.lower():
                break

            await next_btn.click()
            await page.wait_for_timeout(1500)

        attachments = filter_by_date_range(dedupe_attachments(found), args.date_from, args.date_to)
        if args.max_pdfs and args.max_pdfs > 0:
            attachments = attachments[: args.max_pdfs]

        print(f"Found {len(attachments)} target KSEI attachment PDFs")

        cookies = await context.cookies()
        jar = requests.cookies.RequestsCookieJar()
        for c in cookies:
            jar.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))

        sess = requests.Session()
        sess.cookies = jar
        sess.headers.update({"User-Agent": await page.evaluate("() => navigator.userAgent")})

        downloaded: list[Path] = []
        for item in tqdm(attachments, desc="Downloading PDFs"):
            out_path = out_dir / item.attachment_filename
            item.local_pdf_path = str(out_path)

            if (not args.overwrite) and out_path.exists() and out_path.stat().st_size > 0:
                item.download_status = "exists"
                downloaded.append(out_path)
                continue

            time.sleep(0.2)
            try:
                r = sess.get(item.attachment_url, timeout=90)
                r.raise_for_status()
                data = r.content
                if not looks_like_pdf(data):
                    item.download_status = "skipped"
                    item.error = "Response was not a PDF"
                    print(f"[SKIP] Not a real PDF: {item.attachment_url}")
                    continue
                out_path.write_bytes(data)
                item.download_status = "downloaded"
                downloaded.append(out_path)
            except Exception as exc:
                item.download_status = "error"
                item.error = str(exc)
                print(f"[ERROR] {item.attachment_filename}: {exc}")

        await browser.close()

    write_manifest(attachments, manifest_path)
    print(f"Done. PDFs saved to: {out_dir.resolve()}")
    print(f"Manifest: {manifest_path.resolve()}")

    if args.extract and downloaded:
        extract_out_dir = Path(args.extract_out_dir)
        extract_out_dir.mkdir(parents=True, exist_ok=True)
        print(f"[extract] Running ownership extractor on {len(downloaded)} PDFs")

        for pdf_path in tqdm(downloaded, desc="Extracting to XLSX"):
            out_xlsx = extract_out_dir / f"{pdf_path.stem}.ownership_table.xlsx"
            if out_xlsx.exists() and not args.overwrite:
                continue
            try:
                run_extractor(pdf_path, out_xlsx, manifest_path, debug_root=args.extract_debug_dir)
            except subprocess.CalledProcessError as exc:
                print(f"[extract][ERROR] {pdf_path.name}: {exc}")


if __name__ == "__main__":
    parsed = build_argparser().parse_args()
    asyncio.run(main(parsed))
