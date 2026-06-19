from scrape_and_download import looks_like_pdf, safe_filename


def test_safe_filename_uses_and_sanitizes_link_text() -> None:
    assert (
        safe_filename("https://example.com/download", "20260101 report: 5%")
        == "20260101 report_ 5_.pdf"
    )


def test_safe_filename_falls_back_to_url_path() -> None:
    assert safe_filename("https://example.com/files/report.pdf?download=1") == "report.pdf"


def test_pdf_validation_checks_header_and_minimum_size() -> None:
    assert looks_like_pdf(b"%PDF" + b"0" * 1020)
    assert not looks_like_pdf(b"<html>" + b"0" * 1020)
    assert not looks_like_pdf(b"%PDF")
