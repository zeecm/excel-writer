import os
from enum import Enum
from typing import NamedTuple

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
