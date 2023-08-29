import os
from enum import Enum
from typing import List, NamedTuple

from excel_writer.writer import CellRange

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


class Content(NamedTuple):
    title: str
    descriptions: List[str]


FIRST_CONTENT_TITLE_CELL = "B16"
FIRST_CONTENT_DESCRIPTION_CELL = "B17"

SIGNATURE_BLOCK_CELL_RANGE = CellRange(
    start_row=19, start_column=2, end_row=28, end_column=3
)
