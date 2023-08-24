import os
from enum import Enum
from typing import NamedTuple

from excel_writer.writer import ExcelWriter, Writer

TEMPLATE_FILEPATH = os.path.join(
    "acknowledgement_form", "template", "template_job_ack.xlsx"
)


class FieldInfo(NamedTuple):
    cell_id: str
    template_str: str


class Field(Enum):
    CLIENT_NAME = FieldInfo("B3", "<client_name>")
    JOB_NUM = FieldInfo("B4", "<job_num>")
    PO_NUM = FieldInfo("B5", "<po_num>")
    QUOTATION_NUM = FieldInfo("B6", "<quotation_num>")
    VESSEL = FieldInfo("B7", "<vessel>")
    DRAWING_NUM = FieldInfo("B11", "<drawing_num>")
    CLASS = FieldInfo("B12", "<class>")
    DURATION = FieldInfo("B13", "<duration>")


class FieldValue(NamedTuple):
    field: Field
    value: str


def load_template(template_filepath: str = TEMPLATE_FILEPATH) -> ExcelWriter:
    return ExcelWriter(template_filepath)


def set_field_value(
    writer: ExcelWriter, field: Field, value_to_set: str
) -> ExcelWriter:
    if not value_to_set:
        value_to_set = "-"
        if field == Field.CLASS:
            value_to_set = "Not Involved"

    field_cell_id = field.value.cell_id
    field_template_str = field.value.template_str

    client_name_cell = writer.cell(0, field_cell_id)

    template_value = str(client_name_cell.value)

    new_value = template_value.replace(field_template_str, value_to_set)
    writer.cell(0, field_cell_id, new_value)

    return writer


def create_job_ack(
    filename: str, filepath: str = ".", template_filepath: str = TEMPLATE_FILEPATH
) -> None:
    query_mapping = {}
    for field in Field:
        field_value = input(f"{field.name}: ")
        query_mapping |= {field: field_value}

    writer = load_template(template_filepath)

    for field, value in query_mapping.items():
        writer = set_field_value(writer, field, value)

    writer.save_workbook(filepath, filename)


if __name__ == "__main__":
    filename = "test.xls"
    create_job_ack(filename)
