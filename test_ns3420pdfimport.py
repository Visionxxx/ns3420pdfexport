"""
Tests for ns3420pdfimport.py

Run with:
    cd /home/glenn/Dokumenter/Prosjekt/ns3420pdfimport
    python3 -m pytest test_ns3420pdfimport.py -v
"""

import csv
import io
import json
import os
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from ns3420pdfimport import (
    CoordinateParser,
    Document,
    Post,
    export_csv,
    export_json,
    export_xml,
)

# Path to the integration test PDF
PDF_PATH = Path("/home/glenn/Dokumenter/Prosjekt/Rust/ns3420reader/NS 3420 Beskrivelse - Fløy O.pdf")

PDF_EXISTS = PDF_PATH.exists()


# ────────────────────────────────────────────────────────────
# Fixtures
# ────────────────────────────────────────────────────────────

@pytest.fixture
def parser():
    return CoordinateParser()


@pytest.fixture
def sample_doc():
    """A minimal Document with a few posts for export tests."""
    posts = [
        Post(
            post_number="05.21.2",
            ns3420_code="LB1.1112A",
            subject="Grunnarbeider",
            description="Armert betong",
            quantity=320.0,
            unit="m²",
            unit_price=None,
            total_price=None,
        ),
        Post(
            post_number="05.21.3",
            ns3420_code="LB1.2211",
            subject="Grunnarbeider",
            description="Fundament betong klasse B35",
            quantity=45.5,
            unit="m³",
            unit_price=1500.0,
            total_price=68250.0,
        ),
        Post(
            post_number="73.731.130",
            ns3420_code="",
            subject="Vann og avløp",
            description="Geotekstil, type X",
            quantity=0.0,
            unit="RS",
        ),
    ]
    return Document(
        project_name="Testprosjekt",
        document_name="NS 3420 Beskrivelse",
        date="01.01.2025",
        client="Test Eiendom AS",
        author="Test Rådgiver AS",
        posts=posts,
        chapters={"05": "Betongarbeider", "73": "Vann og avløp"},
    )


# ────────────────────────────────────────────────────────────
# CoordinateParser: _col_for_x
# ────────────────────────────────────────────────────────────

class TestColForX:
    """Column classification based on x-coordinate boundaries."""

    # COL_RANGES:
    #   postnr: [39, 99)
    #   spec:   [99, 365)
    #   enh:    [365, 397)
    #   mengde: [397, 460)
    #   pris:   [460, 524)
    #   sum:    [524, 580)

    def test_postnr_at_boundary_start(self, parser):
        assert parser._col_for_x(39.0) == "postnr"

    def test_postnr_at_x44(self, parser):
        assert parser._col_for_x(44.0) == "postnr"

    def test_postnr_just_before_spec(self, parser):
        assert parser._col_for_x(98.9) == "postnr"

    def test_spec_at_boundary_start(self, parser):
        assert parser._col_for_x(99.0) == "spec"

    def test_spec_at_x100(self, parser):
        assert parser._col_for_x(100.0) == "spec"

    def test_spec_middle(self, parser):
        assert parser._col_for_x(200.0) == "spec"

    def test_spec_just_before_enh(self, parser):
        assert parser._col_for_x(364.9) == "spec"

    def test_enh_at_x370(self, parser):
        assert parser._col_for_x(370.0) == "enh"

    def test_enh_at_boundary_start(self, parser):
        assert parser._col_for_x(365.0) == "enh"

    def test_enh_just_before_mengde(self, parser):
        assert parser._col_for_x(396.9) == "enh"

    def test_mengde_at_x410(self, parser):
        assert parser._col_for_x(410.0) == "mengde"

    def test_mengde_at_boundary_start(self, parser):
        assert parser._col_for_x(397.0) == "mengde"

    def test_mengde_just_before_pris(self, parser):
        assert parser._col_for_x(459.9) == "mengde"

    def test_pris_at_boundary_start(self, parser):
        assert parser._col_for_x(460.0) == "pris"

    def test_pris_middle(self, parser):
        assert parser._col_for_x(490.0) == "pris"

    def test_sum_at_boundary_start(self, parser):
        assert parser._col_for_x(524.0) == "sum"

    def test_sum_middle(self, parser):
        assert parser._col_for_x(550.0) == "sum"

    def test_before_postnr_returns_none(self, parser):
        assert parser._col_for_x(38.9) is None

    def test_after_sum_returns_none(self, parser):
        assert parser._col_for_x(580.0) is None

    def test_far_right_returns_none(self, parser):
        assert parser._col_for_x(700.0) is None


# ────────────────────────────────────────────────────────────
# CoordinateParser: _is_continuation
# ────────────────────────────────────────────────────────────

class TestIsContinuation:
    """Post number split-row detection logic."""

    def test_is_continuation_trailing_dot(self, parser):
        """Pattern A: previous ends with '.' -> next is appended."""
        assert parser._is_continuation("09.235.10.", "1.1") is True

    def test_is_continuation_trailing_dot_simple(self, parser):
        assert parser._is_continuation("73.731.13.", "0") is True

    def test_is_continuation_digit_single(self, parser):
        """Pattern B: next is only 1-3 digits -> concatenation."""
        assert parser._is_continuation("73.731.13", "0") is True

    def test_is_continuation_digit_two(self, parser):
        assert parser._is_continuation("73.731.1", "30") is True

    def test_is_continuation_digit_three(self, parser):
        assert parser._is_continuation("05.21.2", "345") is True

    def test_is_continuation_sub_number(self, parser):
        """Pattern C: next is N.M format -> sub-post continuation."""
        assert parser._is_continuation("73.731.19", "8.1") is True

    def test_is_continuation_sub_number_double_digit(self, parser):
        assert parser._is_continuation("09.235.10", "1.2") is True

    def test_not_continuation_full_postnr(self, parser):
        """A full post number like '05.21.3' after '05.21.2' is NOT a continuation."""
        assert parser._is_continuation("05.21.2", "05.21.3") is False

    def test_not_continuation_new_chapter(self, parser):
        assert parser._is_continuation("05.21.2", "73.731.1") is False

    def test_not_continuation_four_digit_next(self, parser):
        """Next with 4+ digits is not Pattern B (only 1-3 digits allowed)."""
        assert parser._is_continuation("05.21.2", "1234") is False

    def test_not_continuation_non_postnr_prev(self, parser):
        """Previous doesn't look like a post number -> not a continuation."""
        assert parser._is_continuation("HEADING", "1") is False


# ────────────────────────────────────────────────────────────
# CoordinateParser: _merge_postnr
# ────────────────────────────────────────────────────────────

def _make_row(postnr="", spec="", enh="", mengde="", pris="", summ="", page=1, ch_code="05", ch_title="Test"):
    """Helper to create a raw row dict."""
    return {
        "postnr": postnr,
        "spec": spec,
        "enh": enh,
        "mengde": mengde,
        "pris": pris,
        "sum": summ,
        "page": page,
        "ch_code": ch_code,
        "ch_title": ch_title,
    }


class TestMergePostnr:
    """Tests for pass-2 post number merging logic."""

    def test_merge_postnr_trailing_dot(self, parser):
        """'09.235.10.' + '1.1' -> '09.235.10.1.1'."""
        rows = [
            _make_row(postnr="09.235.10.", spec="First line"),
            _make_row(postnr="1.1", spec="Second line"),
        ]
        merged = parser._merge_postnr(rows)
        assert len(merged) == 1
        assert merged[0]["postnr"] == "09.235.10.1.1"

    def test_merge_postnr_digit_concat(self, parser):
        """'73.731.13' + '0' -> '73.731.130'."""
        rows = [
            _make_row(postnr="73.731.13", spec="Geotekstil"),
            _make_row(postnr="0", spec="GEOTEKSTIL klasse X"),
        ]
        merged = parser._merge_postnr(rows)
        assert len(merged) == 1
        assert merged[0]["postnr"] == "73.731.130"

    def test_merge_postnr_sub(self, parser):
        """'73.731.19' + '8.1' -> '73.731.198.1'."""
        rows = [
            _make_row(postnr="73.731.19", spec="Rørledning"),
            _make_row(postnr="8.1", spec="DN 100"),
        ]
        merged = parser._merge_postnr(rows)
        assert len(merged) == 1
        assert merged[0]["postnr"] == "73.731.198.1"

    def test_no_merge_standalone_posts(self, parser):
        """Two distinct post numbers should NOT be merged."""
        rows = [
            _make_row(postnr="05.21.2", spec="Post A"),
            _make_row(postnr="05.21.3", spec="Post B"),
        ]
        merged = parser._merge_postnr(rows)
        assert len(merged) == 2
        assert merged[0]["postnr"] == "05.21.2"
        assert merged[1]["postnr"] == "05.21.3"

    def test_merge_spec_appended(self, parser):
        """Spec text from continuation row should be appended."""
        rows = [
            _make_row(postnr="09.235.10.", spec="First line"),
            _make_row(postnr="1.1", spec="Second line"),
        ]
        merged = parser._merge_postnr(rows)
        assert "First line" in merged[0]["spec"]
        assert "Second line" in merged[0]["spec"]

    def test_merge_enh_from_next_row(self, parser):
        """Unit (enh) should be taken from continuation row if first row is missing it."""
        rows = [
            _make_row(postnr="09.235.10.", spec="Betong", enh=""),
            _make_row(postnr="1.1", spec="", enh="m²", mengde="50,00"),
        ]
        merged = parser._merge_postnr(rows)
        assert merged[0]["enh"] == "m²"
        assert merged[0]["mengde"] == "50,00"

    def test_empty_rows_passthrough(self, parser):
        assert parser._merge_postnr([]) == []


# ────────────────────────────────────────────────────────────
# CoordinateParser: _parse_qty
# ────────────────────────────────────────────────────────────

class TestParseQty:
    """Norwegian number string to float conversion."""

    def test_parse_norwegian_decimal(self):
        assert CoordinateParser._parse_qty("350,00") == pytest.approx(350.0)

    def test_parse_integer_string(self):
        assert CoordinateParser._parse_qty("100") == pytest.approx(100.0)

    def test_parse_empty_string(self):
        assert CoordinateParser._parse_qty("") == 0.0

    def test_parse_none_equivalent(self):
        assert CoordinateParser._parse_qty("  ") == 0.0

    def test_parse_with_spaces(self):
        # Some quantities have space as thousands separator
        assert CoordinateParser._parse_qty("1 350,00") == pytest.approx(1350.0)

    def test_parse_dot_decimal(self):
        assert CoordinateParser._parse_qty("123.45") == pytest.approx(123.45)

    def test_parse_invalid_returns_zero(self):
        assert CoordinateParser._parse_qty("n/a") == 0.0

    def test_parse_large_quantity(self):
        assert CoordinateParser._parse_qty("12345,67") == pytest.approx(12345.67)

    def test_parse_zero_string(self):
        assert CoordinateParser._parse_qty("0,00") == pytest.approx(0.0)


# ────────────────────────────────────────────────────────────
# CoordinateParser: _split_ns_code
# ────────────────────────────────────────────────────────────

class TestSplitNsCode:
    """NS code extraction from spec text."""

    def test_extract_standard_ns_code(self, parser):
        code, desc = parser._split_ns_code("LB1.1112A Armert betong")
        assert code == "LB1.1112A"
        assert "Armert betong" in desc

    def test_extract_code_multiline(self, parser):
        spec = "NB2.162A STØPT PÅ STEDET\nType X\nAreal m 320,00"
        code, desc = parser._split_ns_code(spec)
        assert code == "NB2.162A"
        assert "STØPT PÅ STEDET" in desc

    def test_no_ns_code_returns_empty_code(self, parser):
        code, desc = parser._split_ns_code("Noen vanlig tekst uten kode")
        assert code == ""
        assert desc == "Noen vanlig tekst uten kode"

    def test_empty_spec(self, parser):
        code, desc = parser._split_ns_code("")
        assert code == ""
        assert desc == ""

    def test_short_code_wza(self, parser):
        """Short codes like WZA (2-4 uppercase letters + optional digit) should match."""
        code, desc = parser._split_ns_code("WZA Beskrivelse av post")
        assert code == "WZA"

    def test_non_ns_word_not_extracted(self, parser):
        """Regular word starting with uppercase letters should NOT be extracted as code."""
        code, desc = parser._split_ns_code("BETONGARBEIDER klasse B35")
        assert code == ""

    def test_code_with_suffix(self, parser):
        code, desc = parser._split_ns_code("SB1.8221 Stålplate")
        assert code == "SB1.8221"


# ────────────────────────────────────────────────────────────
# CoordinateParser: _normalize_unit
# ────────────────────────────────────────────────────────────

class TestNormalizeUnit:
    """Unit string normalization."""

    def test_m2_to_superscript(self):
        assert CoordinateParser._normalize_unit("m2") == "m²"

    def test_m3_to_superscript(self):
        assert CoordinateParser._normalize_unit("m3") == "m³"

    def test_stk_unchanged(self):
        assert CoordinateParser._normalize_unit("stk") == "stk"

    def test_m_unchanged(self):
        assert CoordinateParser._normalize_unit("m") == "m"

    def test_rs_unchanged(self):
        assert CoordinateParser._normalize_unit("RS") == "RS"

    def test_kg_unchanged(self):
        assert CoordinateParser._normalize_unit("kg") == "kg"

    def test_already_superscript_m2_unchanged(self):
        assert CoordinateParser._normalize_unit("m²") == "m²"

    def test_already_superscript_m3_unchanged(self):
        assert CoordinateParser._normalize_unit("m³") == "m³"

    def test_empty_string(self):
        assert CoordinateParser._normalize_unit("") == ""


# ────────────────────────────────────────────────────────────
# Export: CSV
# ────────────────────────────────────────────────────────────

class TestExportCsv:
    """CSV export correctness."""

    def test_export_csv_creates_file(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        assert out.exists()

    def test_export_csv_header_row(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        lines = content.splitlines()
        header = lines[0]
        assert "Postnr" in header
        assert "NS3420" in header
        assert "Emne" in header
        assert "Beskrivelse" in header
        assert "Mengde" in header
        assert "Enhet" in header

    def test_export_csv_post_count(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        reader = csv.reader(io.StringIO(content), delimiter=";")
        rows = list(reader)
        # 1 header + 3 posts
        assert len(rows) == 4

    def test_export_csv_post_number(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        assert "05.21.2" in content

    def test_export_csv_ns_code(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        assert "LB1.1112A" in content

    def test_export_csv_quantity(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        assert "320.00" in content

    def test_export_csv_unit(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        assert "m²" in content

    def test_export_csv_price_fields(self, sample_doc, tmp_path):
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        # Post with unit_price=1500.0 and total_price=68250.0
        assert "1500.00" in content
        assert "68250.00" in content

    def test_export_csv_empty_quantity_shows_empty(self, sample_doc, tmp_path):
        """Post with quantity=0 should produce empty quantity cell, not '0.00'."""
        out = tmp_path / "test.csv"
        export_csv(sample_doc, out)
        content = out.read_text(encoding="utf-8-sig")
        reader = csv.reader(io.StringIO(content), delimiter=";")
        rows = list(reader)
        # Find the row for post 73.731.130 (quantity=0)
        rs_row = next((r for r in rows if r[0] == "73.731.130"), None)
        assert rs_row is not None
        # Mengde column (index 4) should be empty for quantity=0
        assert rs_row[4] == ""


# ────────────────────────────────────────────────────────────
# Export: JSON
# ────────────────────────────────────────────────────────────

class TestExportJson:
    """JSON export correctness."""

    def test_export_json_creates_file(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        assert out.exists()

    def test_export_json_valid_structure(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert "metadata" in data
        assert "posts" in data

    def test_export_json_metadata(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        meta = data["metadata"]
        assert meta["project_name"] == "Testprosjekt"
        assert meta["document_name"] == "NS 3420 Beskrivelse"
        assert meta["total_posts"] == 3

    def test_export_json_chapters_in_metadata(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert "05" in data["metadata"]["chapters"]
        assert data["metadata"]["chapters"]["05"] == "Betongarbeider"

    def test_export_json_post_count(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert len(data["posts"]) == 3

    def test_export_json_post_fields(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        post = next(p for p in data["posts"] if p["postnr"] == "05.21.2")
        assert post["ns3420"] == "LB1.1112A"
        assert post["mengde"] == pytest.approx(320.0)
        assert post["enhet"] == "m²"

    def test_export_json_price_fields(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        post = next(p for p in data["posts"] if p["postnr"] == "05.21.3")
        assert post["pris"] == pytest.approx(1500.0)
        assert post["sum"] == pytest.approx(68250.0)

    def test_export_json_null_price_is_null(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        post = next(p for p in data["posts"] if p["postnr"] == "05.21.2")
        assert post["pris"] is None
        assert post["sum"] is None

    def test_export_json_norwegian_chars(self, sample_doc, tmp_path):
        out = tmp_path / "test.json"
        export_json(sample_doc, out)
        content = out.read_text(encoding="utf-8")
        # ensure_ascii=False so Norwegian chars should appear directly
        assert "m²" in content or "m\u00b2" in content


# ────────────────────────────────────────────────────────────
# Export: XML
# ────────────────────────────────────────────────────────────

class TestExportXml:
    """XML (NS 3459) export correctness.

    The exported XML uses a default namespace (xmlns="http://www.standard.no/ns3459"),
    so element lookups via ElementTree must include the namespace URI.
    """

    NS = "http://www.standard.no/ns3459"

    def _ns(self, tag: str) -> str:
        return f"{{{self.NS}}}{tag}"

    def test_export_xml_creates_file(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        assert out.exists()

    def test_export_xml_parseable(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        assert root is not None

    def test_export_xml_root_element(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        assert "NS3459" in root.tag

    def test_export_xml_metadata(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        meta = root.find(self._ns("Metadata"))
        assert meta is not None
        prosjekt = meta.find(self._ns("Prosjekt"))
        assert prosjekt is not None
        assert prosjekt.text == "Testprosjekt"

    def test_export_xml_post_count(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        posts = root.findall(self._ns("Post"))
        assert len(posts) == 3

    def test_export_xml_post_number(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        postnrs = [p.findtext(self._ns("Postnr")) for p in root.findall(self._ns("Post"))]
        assert "05.21.2" in postnrs

    def test_export_xml_ns_code(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        post = next(
            p for p in root.findall(self._ns("Post"))
            if p.findtext(self._ns("Postnr")) == "05.21.2"
        )
        kode = post.find(self._ns("Kode"))
        assert kode is not None
        assert kode.findtext(self._ns("ID")) == "LB1.1112A"

    def test_export_xml_prisinfo(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        post = next(
            p for p in root.findall(self._ns("Post"))
            if p.findtext(self._ns("Postnr")) == "05.21.2"
        )
        prisinfo = post.find(self._ns("Prisinfo"))
        assert prisinfo is not None
        assert prisinfo.findtext(self._ns("Mengde")) == "320.0"
        assert prisinfo.findtext(self._ns("Enhet")) == "m²"

    def test_export_xml_each_post_has_uuid(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        ids = [p.findtext(self._ns("ID")) for p in root.findall(self._ns("Post"))]
        # All IDs should be unique and non-empty
        assert len(ids) == len(set(ids))
        assert all(i for i in ids)

    def test_export_xml_post_without_ns_code_has_no_kode(self, sample_doc, tmp_path):
        out = tmp_path / "test.xml"
        export_xml(sample_doc, out)
        tree = ET.parse(str(out))
        root = tree.getroot()
        post = next(
            p for p in root.findall(self._ns("Post"))
            if p.findtext(self._ns("Postnr")) == "73.731.130"
        )
        # No NS code -> no <Kode> element
        assert post.find(self._ns("Kode")) is None


# ────────────────────────────────────────────────────────────
# Integration: full PDF parse
# ────────────────────────────────────────────────────────────

@pytest.mark.skipif(not PDF_EXISTS, reason=f"PDF not found at {PDF_PATH}")
class TestFullPdfParse:
    """Integration tests that parse the real NS 3420 PDF."""

    @pytest.fixture(scope="class")
    def parsed_doc(self):
        coord_parser = CoordinateParser()
        return coord_parser.parse(PDF_PATH)

    def test_minimum_post_count(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        assert len(real_posts) >= 1900, (
            f"Expected at least 1900 posts, got {len(real_posts)}"
        )

    def test_chapter_count(self, parsed_doc):
        assert len(parsed_doc.chapters) >= 40, (
            f"Expected at least 40 chapters, got {len(parsed_doc.chapters)}"
        )

    def test_post_05_21_2_exists(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "05.21.2"), None)
        assert post is not None, "Post '05.21.2' not found"

    def test_post_05_21_2_ns_code(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "05.21.2"), None)
        assert post is not None
        assert post.ns3420_code == "LB1.1112A", (
            f"Expected ns code 'LB1.1112A', got '{post.ns3420_code}'"
        )

    def test_post_05_21_2_unit(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "05.21.2"), None)
        assert post is not None
        assert post.unit in ("m²", "m2"), (
            f"Expected unit 'm²' or 'm2', got '{post.unit}'"
        )

    def test_post_05_21_2_quantity(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "05.21.2"), None)
        assert post is not None
        assert post.quantity == pytest.approx(320.0, abs=1.0), (
            f"Expected quantity ~320.0, got {post.quantity}"
        )

    def test_split_postnr_73_731_130_reassembly(self, parsed_doc):
        """Post 73.731.130 tests that split post-number rows are correctly merged."""
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "73.731.130"), None)
        assert post is not None, (
            "Post '73.731.130' not found; split postnr reassembly may be broken"
        )

    def test_text_overflow_fix_01_19_4_1(self, parsed_doc):
        """Post 01.19.4.1 tests that description text overflowing into postnr column is handled."""
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post = next((p for p in real_posts if p.post_number == "01.19.4.1"), None)
        assert post is not None, (
            "Post '01.19.4.1' not found; text overflow fix may be broken"
        )

    def test_no_duplicate_post_numbers(self, parsed_doc):
        """Almost no duplicate post numbers; allow at most 2 duplicates for known edge cases."""
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        post_numbers = [p.post_number for p in real_posts]
        unique_numbers = set(post_numbers)
        duplicates = [pn for pn in unique_numbers if post_numbers.count(pn) > 1]
        assert len(duplicates) <= 2, f"Too many duplicate post numbers: {duplicates[:10]}"

    def test_document_has_project_name(self, parsed_doc):
        assert parsed_doc.project_name, "Expected a non-empty project name"

    def test_posts_have_chapter_codes(self, parsed_doc):
        real_posts = [p for p in parsed_doc.posts if p.post_number]
        posts_with_chapter = [p for p in real_posts if p.chapter_code]
        # At least 90% of real posts should have a chapter code
        ratio = len(posts_with_chapter) / len(real_posts)
        assert ratio >= 0.9, (
            f"Only {ratio*100:.1f}% of posts have a chapter code"
        )
