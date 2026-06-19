import pytest

from extract_ownership_table import (
    _compact_spaced_text,
    extract_numeric_tail,
    looks_like_continuation,
    looks_like_new_record,
    parse_int_token,
    parse_pct_token,
    resolve_page_window,
)


def test_compacts_spaced_pdf_glyphs() -> None:
    assert _compact_spaced_text("I N D O N E S I A") == "INDONESIA"
    assert _compact_spaced_text("5 9 9 , 4 9 1 , 2 2 8") == "599,491,228"


def test_parses_numeric_tokens() -> None:
    assert parse_int_token("Shares: 1,234,567") == 1_234_567
    assert parse_pct_token("Ownership 4 1 . 1 0") == pytest.approx(41.10)


def test_extracts_six_value_numeric_tail_without_delta() -> None:
    values = extract_numeric_tail(
        "ABCD 1,000 2,000 10.00 1,500 2,500 12.50"
    )

    assert values == {
        "shares_before": 1_000,
        "combined_before": 2_000,
        "pct_before": 10.0,
        "shares_after": 1_500,
        "combined_after": 2_500,
        "pct_after": 12.5,
        "delta_shares_or_delta_field": "",
    }


def test_identifies_new_and_continuation_rows() -> None:
    new_row = {
        "row_no": "12",
        "ticker": "ABCD",
        "status": "L",
        "shares_before": "1,000",
    }
    continuation = {"address": "JAKARTA SELATAN"}

    assert looks_like_new_record(new_row)
    assert looks_like_continuation(continuation, prev_ticker="ABCD")


def test_max_pages_is_counted_from_selected_start_page() -> None:
    assert resolve_page_window(
        total_pages=20,
        table_page_idx=9,
        max_pages=5,
    ) == (9, 14)


def test_rejects_inverted_page_window() -> None:
    with pytest.raises(ValueError, match="end page"):
        resolve_page_window(
            total_pages=20,
            table_page_idx=0,
            page_from=8,
            page_to=3,
        )
