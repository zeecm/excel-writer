from typing import Dict, List

from loguru import logger
from pypdf import PdfReader

from acknowledgement_form.form_generator.constants import Field

CLIENT_NAME_TEXT_COORDINATES = "18.48, 590.065, 217.986, 599.048"

SAMPLE_PDF = "tests/acknowledgement_form_tests/test_files/sample_quo.pdf"


def get_fields_from_quotation_pdf(quotation_pdf_filepath: str) -> Dict[Field, str]:
    reader = PdfReader(quotation_pdf_filepath)
    return _get_field_values_from_pdf_reader(reader)


def _get_field_values_from_pdf_reader(reader: PdfReader) -> Dict[Field, str]:
    all_pages_text = _extract_text_from_pages(reader)

    first_page_text = all_pages_text[0]
    first_page_fields = _get_fields_from_first_page(first_page_text)

    duration = _get_duration_from_pages(all_pages_text)

    return first_page_fields | {Field.DURATION: duration}


def _extract_text_from_pages(reader: PdfReader) -> List[str]:
    return [page.extract_text() for page in reader.pages]


def _get_fields_from_first_page(first_page_text: str) -> Dict[Field, str]:
    client_name = get_client_name(first_page_text)
    quotation_number = get_quotation_number(first_page_text)
    vessel = get_vessel(first_page_text)
    vessel_class = get_vessel_class(first_page_text)
    drawing_number = get_drawing_number(first_page_text)
    return {
        Field.CLIENT_NAME: client_name,
        Field.QUOTATION_NUM: quotation_number,
        Field.VESSEL: vessel,
        Field.CLASS: vessel_class,
        Field.DRAWING_NUM: drawing_number,
    }


def _get_duration_from_pages(all_pages_text: List[str]) -> str:
    if page_with_duration := [
        page_text for page_text in all_pages_text if "Duration:" in page_text
    ]:
        return get_duration(page_with_duration[0])
    return "N/A"


def get_client_name(page_text: str) -> str:
    bill_to_text = "BILL TO\n"

    if bill_to_text not in page_text:
        logger.warning("Client name could not be found")
        return "-"

    bill_to_text_index = page_text.find(bill_to_text)

    client_name_start_index = bill_to_text_index + len(bill_to_text)
    client_name_end_index = page_text.find("\n", client_name_start_index)

    return page_text[client_name_start_index:client_name_end_index].strip()


def get_quotation_number(page_text: str) -> str:
    quotation_prefix = "MMSQ"

    if quotation_prefix not in page_text:
        logger.warning("Quotation number could not be found")
        return "-"

    quotation_number_start_index = page_text.find(quotation_prefix)
    quotation_number_end_index = page_text.find(" ", quotation_number_start_index)

    # in the parsed pdf text, version number is infront of page no. text
    version_number_prefix = "Page No. "
    version_number_index = page_text.find(version_number_prefix) + len(
        version_number_prefix
    )
    version_number = int(page_text[version_number_index : version_number_index + 1])

    version_number_str = f" V{version_number}" if version_number > 1 else ""
    quotation_number = page_text[
        quotation_number_start_index:quotation_number_end_index
    ]
    return (quotation_number + version_number_str).strip()


def get_vessel(page_text: str) -> str:
    vessel_prefix = "Vessel: "

    if vessel_prefix not in page_text:
        logger.warning("Vessel could not be found")
        return "N/A"

    vessel_start_index = page_text.find(vessel_prefix) + len(vessel_prefix)
    vessel_end_index = page_text.find("\n", vessel_start_index)
    return page_text[vessel_start_index:vessel_end_index].strip()


def get_vessel_class(page_text: str) -> str:
    class_prefix = "Class: "

    if class_prefix not in page_text:
        logger.warning("Class could not be found")
        return "Not Involved"

    class_start_index = page_text.find(class_prefix) + len(class_prefix)
    class_end_index = page_text.find("\n", class_start_index)
    return page_text[class_start_index:class_end_index].strip()


def get_duration(page_text: str) -> str:
    duration_prefix = "Duration:\n"

    if duration_prefix not in page_text:
        logger.warning("Duration could not be found")
        return "N/A"

    duration_start_index = page_text.find(duration_prefix) + len(duration_prefix)
    duration_end_index = _get_end_index_for_duration(page_text, duration_start_index)
    return page_text[duration_start_index:duration_end_index].strip()


def _get_end_index_for_duration(page_text: str, start_index) -> int:
    possible_end_strs = ["day", "month", "week"]
    for end_str in possible_end_strs:
        end_str_index = page_text.lower().find(end_str, start_index)
        if end_str_index != -1:
            index = end_str_index + len(end_str)
            if page_text[index].lower() == "s":
                index = index + 1
            return index
    alterative_duration_end_prefix = "Delivery:"
    return page_text.find(alterative_duration_end_prefix, start_index)


def get_drawing_number(page_text: str) -> str:
    drawing_number_prefix = "Drawing No.:"

    if drawing_number_prefix not in page_text:
        logger.warning("no drawing number found")
        return "-"

    drawing_number_start_index = page_text.find(drawing_number_prefix) + len(
        drawing_number_prefix
    )
    drawing_number_end_index = page_text.find("\n", drawing_number_start_index)
    return page_text[drawing_number_start_index:drawing_number_end_index].strip()