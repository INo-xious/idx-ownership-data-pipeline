from pathlib import Path

import pytest

from extract_ownership_table import (
    _compact_spaced_text,
    _norm,
    compute_change_fields,
    extract_numeric_tail,
    extract_pdf_to_frames,
    extract_report_dates_from_text,
    highlighted_columns,
    highlighted_text,
    is_blue_font_color,
    parse_pdf_date_token,
    parse_pct_token,
)
from scrape_and_download import (
    DisclosureAttachment,
    dedupe_attachments,
    filename_date,
    is_target_attachment_filename,
    normalize_disclosure_candidate,
    parse_indonesian_datetime,
)


FIXTURE_DIR = Path(__file__).parent / "fixtures"
SAMPLE_PDF = Path("/Users/marvelharisson/Downloads/20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf")


def test_normalizes_repeated_whitespace() -> None:
    assert _norm("  PT   SAMPLE\nDATA  ") == "PT SAMPLE DATA"


def test_compacts_spaced_country_name() -> None:
    assert _compact_spaced_text("I N D O N E S I A") == "INDONESIA"


def test_compacts_spaced_share_count() -> None:
    assert _compact_spaced_text("5 9 9 , 4 9 1 , 2 2 8") == "599,491,228"


def test_preserves_ambiguous_short_number_sequence() -> None:
    assert _compact_spaced_text("1 . 2") == "1 . 2"


def test_parses_standard_percentage() -> None:
    assert parse_pct_token("41.10") == pytest.approx(41.10)


def test_parses_spaced_percentage() -> None:
    assert parse_pct_token("4 1 . 1 0") == pytest.approx(41.10)


def test_parses_percentage_inside_labelled_text() -> None:
    assert parse_pct_token("Ownership after: 7.25 percent") == pytest.approx(7.25)


def test_rejects_percentage_without_two_decimal_places() -> None:
    assert parse_pct_token("Ownership: 7.2 percent") is None


def test_extracts_numeric_tail_from_standard_fixture() -> None:
    text = (FIXTURE_DIR / "ownership_row_standard.txt").read_text(encoding="utf-8")

    assert extract_numeric_tail(text) == {
        "shares_before": 1_000_000,
        "combined_before": 1_200_000,
        "pct_before": 10.0,
        "shares_after": 1_100_000,
        "combined_after": 1_300_000,
        "pct_after": 11.0,
        "delta_shares_or_delta_field": "100,000",
    }


def test_extracts_numeric_tail_from_spaced_fixture() -> None:
    text = (FIXTURE_DIR / "ownership_row_spaced.txt").read_text(encoding="utf-8")

    assert extract_numeric_tail(text) == {
        "shares_before": 200_000,
        "combined_before": 300_000,
        "pct_before": 20.0,
        "shares_after": 250_000,
        "combined_after": 350_000,
        "pct_after": 25.0,
        "delta_shares_or_delta_field": "50,000",
    }


def test_parses_indonesian_disclosure_datetime() -> None:
    assert parse_indonesian_datetime("08 Juni 2026 17:37:20") == "2026-06-08 17:37:20"


def test_filters_target_idx_attachment_filenames() -> None:
    assert is_target_attachment_filename("20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf")
    assert not is_target_attachment_filename("Pemegang Saham di atas 5% (KSEI) [Semua Emiten Saham].pdf")
    assert filename_date("20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf") == "2026-06-08"


def test_normalizes_disclosure_candidate_metadata() -> None:
    raw = {
        "href": "https://www.idx.co.id/StaticData/NewsAndAnnouncement/ANNOUNCEMENTSTOCK/From_EREP/20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf",
        "text": "20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf",
        "container_text": (
            "08 Juni 2026 17:37:20\n"
            "Pemegang Saham di atas 5% (KSEI) [Semua Emiten Saham]\n"
            "@ 20260608_Semua Emiten Saham_Pengumuman Bursa_32098373_lamp1.pdf"
        ),
        "detail_url": "https://www.idx.co.id/id/perusahaan-tercatat/keterbukaan-informasi/detail",
    }

    item = normalize_disclosure_candidate(raw)

    assert item is not None
    assert item.disclosure_datetime == "2026-06-08 17:37:20"
    assert item.announcement_date == "2026-06-08"
    assert item.filename_date == "2026-06-08"
    assert "Pemegang Saham di atas 5%" in item.disclosure_title


def test_deduplicates_attachments_preferring_richer_metadata() -> None:
    weak = DisclosureAttachment("", "", "", "", "20260608_a_lamp1.pdf", "https://idx/a.pdf", "2026-06-08")
    rich = DisclosureAttachment(
        "Pemegang Saham di atas 5% (KSEI)",
        "2026-06-08 17:37:20",
        "2026-06-08",
        "https://idx/detail",
        "20260608_a_lamp1.pdf",
        "https://idx/a.pdf",
        "2026-06-08",
    )

    assert dedupe_attachments([weak, rich]) == [rich]


def test_parses_pdf_report_dates() -> None:
    dates = extract_report_dates_from_text(
        "KEPEMILIKAN EFEK DIATAS 5% BERDASARKAN SID (PUBLIK) per tanggal 5 Jun 2026 "
        "Kepemilikan Per Investor 04-JUN-2026 Kepemilikan Per Investor 05-JUN-2026"
    )

    assert parse_pdf_date_token("05-JUN-2026") == "2026-06-05"
    assert parse_pdf_date_token("5 Jun 2026") == "2026-06-05"
    assert dates == {"report_date": "2026-06-05", "previous_report_date": "2026-06-04"}


def test_detects_blue_highlighted_words_and_columns() -> None:
    words = [
        {"text": "OWNER", "x0": 1, "non_stroking_color": (0.0, 0.0, 1.0)},
        {"text": "BLACK", "x0": 2, "non_stroking_color": (0.0, 0.0, 0.0)},
    ]
    buckets = {"shareholder": [words[0]], "ticker": [words[1]]}

    assert is_blue_font_color((0.0, 0.0, 1.0))
    assert not is_blue_font_color((0.0, 0.0, 0.0))
    assert highlighted_text(words) == "OWNER"
    assert highlighted_columns(buckets) == "shareholder"


def test_computes_change_fields_from_before_after_values() -> None:
    record = {
        "shares_before": 1_000_000,
        "shares_after": 1_250_000,
        "pct_before": 5.0,
        "pct_after": 6.25,
        "is_highlighted_change": True,
    }

    compute_change_fields(record)

    assert record["delta_shares"] == 250_000
    assert record["delta_pct"] == pytest.approx(1.25)
    assert record["has_numeric_change"] is True
    assert record["change_reason"] == "blue_text,numeric_delta"


@pytest.mark.skipif(not SAMPLE_PDF.exists(), reason="sample IDX PDF is not available on this machine")
def test_sample_pdf_smoke_detects_highlighted_changes() -> None:
    ownership_all, ownership_changes, _warnings, _raw = extract_pdf_to_frames(
        SAMPLE_PDF,
        page_from=3,
        page_to=3,
    )

    assert not ownership_all.empty
    assert not ownership_changes.empty
    assert ownership_all["report_date"].dropna().iloc[0] == "2026-06-05"
    assert ownership_changes["is_highlighted_change"].any()
