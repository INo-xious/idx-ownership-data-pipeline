from __future__ import annotations

import os
import json
import shutil
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st


ROOT = Path(__file__).parent

SCRAPE_SCRIPT = ROOT / "scrape_and_download.py"
EXTRACT_SCRIPT = ROOT / "extract_ownership_table.py"

PDF_DIR = ROOT / "outputs" / "pdfs"
EXTRACTED_DIR = ROOT / "outputs" / "extracted"
MANIFEST_FILE = PDF_DIR / "download_manifest.json"
FINAL_FILE = EXTRACTED_DIR / "ownership_table.xlsx"
DEFAULT_KEYWORD = "Pemegang Saham di atas 5% (KSEI)"


st.set_page_config(page_title="IDX/KSEI Ownership Changes", layout="wide")


def apply_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            --idx-red: #b4232a;
            --idx-red-dark: #811b22;
            --ink: #1f2933;
            --muted: #697586;
            --line: #e4e7ec;
            --soft: #f8fafc;
            --teal: #0e9384;
            --blue: #1d4ed8;
            --gold: #b54708;
        }
        .main .block-container {
            padding-top: 1.4rem;
            padding-bottom: 2.5rem;
            max-width: 1480px;
        }
        h1, h2, h3 {
            color: var(--ink);
            letter-spacing: 0;
        }
        h1 {
            font-size: 2.05rem;
            font-weight: 760;
            margin-bottom: 0.1rem;
        }
        h2 {
            font-size: 1.25rem;
            font-weight: 720;
        }
        h3 {
            font-size: 1.02rem;
            font-weight: 720;
        }
        [data-testid="stSidebar"] {
            background: #ffffff;
            border-right: 1px solid var(--line);
        }
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 0.9rem 1rem;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
        }
        [data-testid="stMetric"] label {
            color: var(--muted);
            font-size: 0.78rem;
        }
        [data-testid="stMetricValue"] {
            color: var(--ink);
            font-size: 1.5rem;
            font-weight: 740;
        }
        .section-panel {
            border: 1px solid var(--line);
            border-radius: 8px;
            background: #ffffff;
            padding: 1rem;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
        }
        .status-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            border-radius: 999px;
            padding: 0.25rem 0.6rem;
            font-size: 0.78rem;
            font-weight: 680;
            border: 1px solid var(--line);
            color: var(--muted);
            background: var(--soft);
        }
        .status-chip::before {
            content: "";
            width: 0.45rem;
            height: 0.45rem;
            border-radius: 999px;
            background: var(--teal);
        }
        .owner-title {
            color: var(--ink);
            font-size: 1.05rem;
            font-weight: 750;
            margin-bottom: 0.2rem;
        }
        .owner-subtitle {
            color: var(--muted);
            font-size: 0.86rem;
            margin-bottom: 0.8rem;
        }
        .detail-row {
            display: flex;
            justify-content: space-between;
            gap: 1rem;
            padding: 0.45rem 0;
            border-bottom: 1px solid #eef2f6;
            font-size: 0.86rem;
        }
        .detail-row span:first-child {
            color: var(--muted);
        }
        .detail-row span:last-child {
            color: var(--ink);
            font-weight: 650;
            text-align: right;
        }
        .mini-metric-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.7rem;
            margin: 0.8rem 0 0.7rem;
        }
        .mini-metric {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 0.7rem;
            min-width: 0;
            background: #ffffff;
        }
        .mini-metric span {
            display: block;
            color: var(--muted);
            font-size: 0.72rem;
            line-height: 1.2;
            margin-bottom: 0.35rem;
        }
        .mini-metric strong {
            display: block;
            color: var(--ink);
            font-size: 1rem;
            line-height: 1.15;
            overflow-wrap: anywhere;
        }
        .stButton > button {
            border-radius: 7px;
            border: 1px solid var(--idx-red);
            background: var(--idx-red);
            color: #ffffff;
            font-weight: 720;
            min-height: 2.55rem;
        }
        .stButton > button:hover {
            border-color: var(--idx-red-dark);
            background: var(--idx-red-dark);
            color: #ffffff;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def run_script(script_path: Path, args: list[str] | None = None) -> str:
    if not script_path.exists():
        raise FileNotFoundError(f"Missing script: {script_path.name}")
    cmd = [sys.executable, str(script_path)]
    if args:
        cmd += args
    proc = subprocess.run(cmd, capture_output=True, text=True)
    output = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    if proc.returncode != 0:
        raise RuntimeError(f"{script_path.name} failed with exit code {proc.returncode}\n{output[-5000:]}")
    return output


def open_folder_in_os(folder: Path) -> None:
    try:
        folder = folder.resolve()
        if sys.platform.startswith("win"):
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        pass


def extract_one_pdf(pdf_path: Path, out_xlsx: Path) -> tuple[bool, str]:
    cmd = [
        sys.executable,
        str(EXTRACT_SCRIPT),
        "--pdf",
        str(pdf_path),
        "--out",
        str(out_xlsx),
        "--manifest",
        str(MANIFEST_FILE),
    ]
    try:
        proc = subprocess.run(cmd, check=True, capture_output=True, text=True)
        return True, (proc.stdout or "")[-2500:]
    except subprocess.CalledProcessError as exc:
        msg = (exc.stdout or "") + "\n" + (exc.stderr or "")
        return False, msg[-5000:]


def read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def load_workbook_frames(path: Path = FINAL_FILE) -> dict[str, pd.DataFrame]:
    return {
        "changes": read_sheet(path, "ownership_changes"),
        "all": read_sheet(path, "ownership_all"),
        "downloads": read_sheet(path, "downloads"),
        "warnings": read_sheet(path, "parse_warnings"),
    }


def merge_workbooks(excel_files: list[Path], out_file: Path) -> tuple[int, int, int]:
    grouped: dict[str, list[pd.DataFrame]] = {
        "ownership_changes": [],
        "ownership_all": [],
        "downloads": [],
        "parse_warnings": [],
    }
    for file in excel_files:
        if file.name.startswith("~$"):
            continue
        for sheet in grouped:
            df = read_sheet(file, sheet)
            if not df.empty:
                df["_source_xlsx"] = file.name
                grouped[sheet].append(df)

    if not grouped["ownership_all"]:
        raise RuntimeError("No readable ownership_all sheets to merge.")

    out_file.parent.mkdir(parents=True, exist_ok=True)
    merged = {sheet: pd.concat(frames, ignore_index=True) if frames else pd.DataFrame() for sheet, frames in grouped.items()}
    with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
        for sheet, df in merged.items():
            df.to_excel(xw, index=False, sheet_name=sheet)
    return (
        len(excel_files),
        int(len(merged["ownership_all"])),
        int(len(merged["ownership_changes"])),
    )


def pdfs_from_manifest() -> list[Path]:
    if not MANIFEST_FILE.exists():
        return sorted(PDF_DIR.glob("*.pdf"))
    try:
        data = json.loads(MANIFEST_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return sorted(PDF_DIR.glob("*.pdf"))
    pdfs: list[Path] = []
    for row in data if isinstance(data, list) else []:
        if str(row.get("download_status") or "") in {"error", "skipped"}:
            continue
        path = Path(str(row.get("local_pdf_path") or ""))
        if path.exists() and path.suffix.lower() == ".pdf":
            pdfs.append(path)
    return sorted(pdfs)


def cleanup_outputs() -> None:
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    EXTRACTED_DIR.mkdir(parents=True, exist_ok=True)
    for path in PDF_DIR.glob("*"):
        try:
            if path.is_file():
                path.unlink()
            else:
                shutil.rmtree(path)
        except Exception:
            pass
    for path in EXTRACTED_DIR.glob("*.xlsx"):
        if path.name != FINAL_FILE.name:
            try:
                path.unlink()
            except Exception:
                pass


def fmt_int(value: object) -> str:
    if pd.isna(value):
        return "-"
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value)


def fmt_pct(value: object) -> str:
    if pd.isna(value):
        return "-"
    try:
        return f"{float(value):.2f}%"
    except (TypeError, ValueError):
        return str(value)


def fmt_text(value: object) -> str:
    if value is None or pd.isna(value):
        return "-"
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return "-"
    return text


def apply_filters(df: pd.DataFrame, ticker: str, owner: str, direction: str, warning_only: bool) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    if ticker:
        out = out[out["ticker"].fillna("").astype(str).str.contains(ticker, case=False, na=False)]
    if owner:
        mask = (
            out["shareholder"].fillna("").astype(str).str.contains(owner, case=False, na=False)
            | out["account_name"].fillna("").astype(str).str.contains(owner, case=False, na=False)
            | out["emiten"].fillna("").astype(str).str.contains(owner, case=False, na=False)
        )
        out = out[mask]
    if direction == "Increase":
        out = out[pd.to_numeric(out["delta_shares"], errors="coerce").fillna(0) > 0]
    elif direction == "Decrease":
        out = out[pd.to_numeric(out["delta_shares"], errors="coerce").fillna(0) < 0]
    if warning_only:
        out = out[out["parse_warnings"].fillna("").astype(str).str.len() > 0]
    return out


def detail_rows(row: pd.Series) -> str:
    pairs = [
        ("Disclosure time", row.get("disclosure_datetime", "")),
        ("Announcement date", row.get("announcement_date", "")),
        ("Report date", row.get("report_date", "")),
        ("Previous date", row.get("previous_report_date", "")),
        ("Source PDF", row.get("source_pdf", "")),
        ("Page", row.get("page", "")),
        ("Confidence", fmt_pct(float(row.get("confidence", 0)) * 100 if not pd.isna(row.get("confidence", None)) else None)),
        ("Warnings", row.get("parse_warnings", "") or "-"),
    ]
    return "\n".join(
        f'<div class="detail-row"><span>{label}</span><span>{fmt_text(value)}</span></div>'
        for label, value in pairs
    )


def render_change_table(df: pd.DataFrame) -> None:
    display_cols = [
        "announcement_date",
        "report_date",
        "ticker",
        "emiten",
        "shareholder",
        "shares_before",
        "shares_after",
        "delta_shares",
        "pct_before",
        "pct_after",
        "delta_pct",
        "change_reason",
    ]
    cols = [c for c in display_cols if c in df.columns]
    st.dataframe(
        df[cols],
        width="stretch",
        hide_index=True,
        column_config={
            "announcement_date": st.column_config.TextColumn("Date"),
            "report_date": st.column_config.TextColumn("Report"),
            "ticker": st.column_config.TextColumn("Ticker"),
            "emiten": st.column_config.TextColumn("Emiten"),
            "shareholder": st.column_config.TextColumn("Owner"),
            "shares_before": st.column_config.NumberColumn("Before", format="%d"),
            "shares_after": st.column_config.NumberColumn("After", format="%d"),
            "delta_shares": st.column_config.NumberColumn("Delta", format="%d"),
            "pct_before": st.column_config.NumberColumn("Before %", format="%.2f"),
            "pct_after": st.column_config.NumberColumn("After %", format="%.2f"),
            "delta_pct": st.column_config.NumberColumn("Delta %", format="%.2f"),
            "change_reason": st.column_config.TextColumn("Signal"),
        },
    )


def render_owner_detail(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.subheader("Selected Owner")
    if df.empty:
        st.caption("Run the pipeline or adjust filters to inspect an ownership change.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    labels = [
        f"{idx}: {row.get('ticker', '-') or '-'} - {row.get('shareholder', '-') or row.get('account_name', '-')}"
        for idx, row in df.head(250).iterrows()
    ]
    selected = st.selectbox("Change row", labels, label_visibility="collapsed")
    selected_idx = int(selected.split(":", 1)[0])
    row = df.loc[selected_idx]
    owner = row.get("shareholder", "") or row.get("account_name", "") or "-"
    st.markdown(f'<div class="owner-title">{owner}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="owner-subtitle">{row.get("ticker", "-")} - {row.get("emiten", "-")}</div>',
        unsafe_allow_html=True,
    )

    metric_items = [
        ("Before", fmt_int(row.get("shares_before"))),
        ("After", fmt_int(row.get("shares_after"))),
        ("Delta", fmt_int(row.get("delta_shares"))),
        ("Before %", fmt_pct(row.get("pct_before"))),
        ("After %", fmt_pct(row.get("pct_after"))),
        ("Delta %", fmt_pct(row.get("delta_pct"))),
    ]
    st.markdown(
        '<div class="mini-metric-grid">'
        + "".join(f'<div class="mini-metric"><span>{label}</span><strong>{value}</strong></div>' for label, value in metric_items)
        + "</div>",
        unsafe_allow_html=True,
    )
    st.markdown(detail_rows(row), unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def run_pipeline(date_from: date, date_to: date, max_pdfs: int, overwrite: bool, cleanup_after: bool) -> list[str]:
    logs: list[str] = []
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    EXTRACTED_DIR.mkdir(parents=True, exist_ok=True)

    scrape_args = [
        "--keyword",
        DEFAULT_KEYWORD,
        "--out-pdf-dir",
        str(PDF_DIR),
        "--manifest-out",
        str(MANIFEST_FILE),
        "--date-from",
        date_from.isoformat(),
        "--date-to",
        date_to.isoformat(),
    ]
    if max_pdfs > 0:
        scrape_args += ["--max-pdfs", str(max_pdfs)]
    if overwrite:
        scrape_args.append("--overwrite")

    logs.append(run_script(SCRAPE_SCRIPT, scrape_args))

    pdfs = pdfs_from_manifest()
    if not pdfs:
        raise RuntimeError(f"No PDFs found in {PDF_DIR}")

    per_pdf_outputs: list[Path] = []
    for pdf in pdfs:
        out_xlsx = EXTRACTED_DIR / f"{pdf.stem}.ownership_table.xlsx"
        if out_xlsx.exists() and not overwrite:
            per_pdf_outputs.append(out_xlsx)
            continue
        ok, log = extract_one_pdf(pdf, out_xlsx)
        logs.append(f"--- {pdf.name} ---\n{log}")
        if ok and out_xlsx.exists():
            per_pdf_outputs.append(out_xlsx)

    file_count, all_rows, change_rows = merge_workbooks(per_pdf_outputs, FINAL_FILE)
    logs.append(f"Merged {file_count} files, {all_rows:,} owner rows, {change_rows:,} change rows.")
    if cleanup_after:
        cleanup_outputs()
        logs.append("Cleanup complete. Kept ownership_table.xlsx.")
    return logs


apply_theme()

today = date.today()
default_from = today - timedelta(days=14)

with st.sidebar:
    st.header("Pipeline")
    st.text_input("IDX keyword", value=DEFAULT_KEYWORD, disabled=True)
    picked_dates = st.date_input("Disclosure date range", value=(default_from, today))
    if isinstance(picked_dates, tuple) and len(picked_dates) == 2:
        date_from, date_to = picked_dates
    else:
        date_from, date_to = default_from, today
    max_pdfs = st.number_input("Max PDFs", min_value=0, value=0, step=1, help="0 downloads every matching attachment.")
    overwrite = st.toggle("Overwrite existing outputs", value=False)
    cleanup_after = st.toggle("Cleanup per-PDF files after merge", value=False)
    run_clicked = st.button("Run IDX Pipeline", type="primary", width="stretch")
    st.divider()
    if st.button("Open Output Folder", width="stretch"):
        open_folder_in_os(EXTRACTED_DIR)
    if FINAL_FILE.exists():
        with FINAL_FILE.open("rb") as f:
            st.download_button(
                "Download Master Workbook",
                data=f,
                file_name="ownership_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

if "pipeline_logs" not in st.session_state:
    st.session_state["pipeline_logs"] = []

if run_clicked:
    with st.status("Running IDX collection and extraction", expanded=True) as status:
        try:
            logs = run_pipeline(date_from, date_to, int(max_pdfs), overwrite, cleanup_after)
            st.session_state["pipeline_logs"] = logs
            status.update(label="Pipeline complete", state="complete", expanded=False)
        except Exception as exc:
            st.session_state["pipeline_logs"] = [str(exc)]
            status.update(label="Pipeline failed", state="error", expanded=True)
            st.error(str(exc))

frames = load_workbook_frames()
changes = frames["changes"]
all_rows = frames["all"]
downloads = frames["downloads"]
warnings = frames["warnings"]

st.title("IDX/KSEI Ownership Changes")
st.caption("Keterbukaan Informasi attachments parsed into owner-level changes and disclosure metadata.")
st.markdown('<span class="status-chip">Workbook ready</span>' if FINAL_FILE.exists() else '<span class="status-chip">No workbook yet</span>', unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Disclosures", fmt_int(len(downloads)))
k2.metric("PDFs Downloaded", fmt_int(int(downloads["local_pdf_path"].fillna("").astype(str).str.len().gt(0).sum()) if not downloads.empty and "local_pdf_path" in downloads else 0))
k3.metric("Changed Rows", fmt_int(len(changes)))
k4.metric("Parse Warnings", fmt_int(len(warnings)))

tab_changes, tab_all, tab_downloads, tab_logs = st.tabs(["Changes", "All Owners", "Downloads", "Logs"])

with tab_changes:
    filter_col, detail_col = st.columns([2.2, 1], gap="large")
    with filter_col:
        st.subheader("Ownership Changes")
        f1, f2, f3, f4 = st.columns([1, 1.3, 1, 1])
        ticker_filter = f1.text_input("Ticker", placeholder="e.g. BEEF")
        owner_filter = f2.text_input("Owner / emiten", placeholder="Search owner, account, or company")
        direction_filter = f3.selectbox("Direction", ["All", "Increase", "Decrease"])
        warning_only = f4.toggle("Warnings only", value=False)
        filtered = apply_filters(changes, ticker_filter, owner_filter, direction_filter, warning_only)
        render_change_table(filtered)
    with detail_col:
        render_owner_detail(filtered)

with tab_all:
    st.subheader("All Owners")
    if all_rows.empty:
        st.info("No extracted owner rows yet.")
    else:
        all_filter = st.text_input("Search all owner rows", placeholder="Ticker, owner, or emiten")
        shown = all_rows
        if all_filter:
            mask = (
                shown["ticker"].fillna("").astype(str).str.contains(all_filter, case=False, na=False)
                | shown["shareholder"].fillna("").astype(str).str.contains(all_filter, case=False, na=False)
                | shown["emiten"].fillna("").astype(str).str.contains(all_filter, case=False, na=False)
            )
            shown = shown[mask]
        st.dataframe(shown, width="stretch", hide_index=True)

with tab_downloads:
    st.subheader("Downloads")
    if downloads.empty:
        st.info("No download manifest loaded.")
    else:
        st.dataframe(downloads, width="stretch", hide_index=True)

with tab_logs:
    st.subheader("Pipeline Logs")
    if not st.session_state["pipeline_logs"]:
        st.caption("Run the pipeline to see scraper and extractor logs.")
    for item in st.session_state["pipeline_logs"]:
        st.code(item[-5000:], language="text")
