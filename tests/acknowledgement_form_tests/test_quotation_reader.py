from typing import Dict

import pytest
from pypdf import PdfReader

from acknowledgement_form.form_generator.constants import Field
from acknowledgement_form.form_generator.quotation_reader import (
    QuotationReader,
    get_client_name,
    get_drawing_number,
    get_duration,
    get_quotation_number,
    get_vessel,
    get_vessel_class,
)


@pytest.mark.parametrize(
    "pdf_location, expected_client_name",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "SCHOTTEL FAR EAST (PTE) LTD",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "BRUNTON'S PROPELLERS LTD",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "NOV RIG SOLUTIONS PTE LTD",
        ),
    ],
)
def test_get_client_name(pdf_location: str, expected_client_name: str):
    reader = PdfReader(pdf_location)
    first_page = reader.pages[0]
    first_page_text = first_page.extract_text()
    assert get_client_name(first_page_text) == expected_client_name


@pytest.mark.parametrize(
    "pdf_location, expected_quotation_number",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "MMSQ23-00558",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "MMSQ23-00584",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "MMSQ23-00553 V3",
        ),
    ],
)
def test_get_quotation_number(pdf_location: str, expected_quotation_number: str):
    reader = PdfReader(pdf_location)
    first_page = reader.pages[0]
    first_page_text = first_page.extract_text()
    assert get_quotation_number(first_page_text) == expected_quotation_number


@pytest.mark.parametrize(
    "pdf_location, expected_vessel",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "N/A",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "N/A",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "Valaris 106",
        ),
    ],
)
def test_get_vessel(pdf_location: str, expected_vessel: str):
    reader = PdfReader(pdf_location)
    first_page = reader.pages[0]
    first_page_text = first_page.extract_text()
    assert get_vessel(first_page_text) == expected_vessel


@pytest.mark.parametrize(
    "pdf_location, expected_class",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "Not Involved",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "BV",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "Not Involved",
        ),
    ],
)
def test_get_vessel_class(pdf_location: str, expected_class: str):
    reader = PdfReader(pdf_location)
    first_page = reader.pages[0]
    first_page_text = first_page.extract_text()
    assert get_vessel_class(first_page_text) == expected_class


@pytest.mark.parametrize(
    "pdf_location, expected_duration",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "4-8 Weeks",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "5-6 months",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "5-6 working days",
        ),
    ],
)
def test_get_duration(pdf_location: str, expected_duration: str):
    reader = PdfReader(pdf_location)
    duration_page_text = [
        page.extract_text()
        for page in reader.pages
        if "Duration" in page.extract_text()
    ][0]
    assert get_duration(duration_page_text) == expected_duration


@pytest.mark.parametrize(
    "pdf_location, expected_drawing_number",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            "-",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            "-",
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            "-",
        ),
    ],
)
def test_get_drawing_number(pdf_location: str, expected_drawing_number: str):
    reader = PdfReader(pdf_location)
    first_page = reader.pages[0]
    first_page_text = first_page.extract_text()
    assert get_drawing_number(first_page_text) == expected_drawing_number


@pytest.mark.parametrize(
    "pdf_location, expected_fields",
    [
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo.pdf",
            {
                Field.CLIENT_NAME: "SCHOTTEL FAR EAST (PTE) LTD",
                Field.QUOTATION_NUM: "MMSQ23-00558",
                Field.VESSEL: "N/A",
                Field.CLASS: "Not Involved",
                Field.DRAWING_NUM: "-",
                Field.DURATION: "4-8 Weeks",
            },
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_2.pdf",
            {
                Field.CLIENT_NAME: "BRUNTON'S PROPELLERS LTD",
                Field.QUOTATION_NUM: "MMSQ23-00584",
                Field.VESSEL: "N/A",
                Field.CLASS: "BV",
                Field.DRAWING_NUM: "-",
                Field.DURATION: "5-6 months",
            },
        ),
        (
            "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf",
            {
                Field.CLIENT_NAME: "NOV RIG SOLUTIONS PTE LTD",
                Field.QUOTATION_NUM: "MMSQ23-00553 V3",
                Field.VESSEL: "Valaris 106",
                Field.CLASS: "Not Involved",
                Field.DRAWING_NUM: "-",
                Field.DURATION: "5-6 working days",
            },
        ),
    ],
)
def test_get_fields_from_quotation(
    pdf_location: str, expected_fields: Dict[Field, str]
):
    quotation_reader = QuotationReader(pdf_location)
    assert quotation_reader.get_fields() == expected_fields
