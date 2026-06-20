from pathlib import Path

import pytest

from extract_ownership_table import (
    _compact_spaced_text,
    _norm,
    extract_numeric_tail,
    parse_pct_token,
)


FIXTURE_DIR = Path(__file__).parent / "fixtures"


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
