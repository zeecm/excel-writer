import os

import pytest

from acknowledgement_form.constants import Field
from acknowledgement_form.generator import set_field_value
from excel_writer.writer import ExcelWriter

TEST_FILEPATH = os.path.join(
    "tests", "acknowledgement_form_tests", "test_files", "template_job_ack.xlsx"
)


@pytest.mark.parametrize(
    "field, value_to_set",
    [
        (Field.CLIENT_NAME, "abc pte ltd"),
        (Field.PO_NUM, "123456"),
        (Field.JOB_NUM, "2308001"),
        (Field.QUOTATION_NUM, "MMSQ23-001"),
        (Field.CLASS, "LR"),
        (Field.VESSEL, "rss hello"),
        (Field.DURATION, "2 days"),
    ],
)
def test_set_field_value(field: Field, value_to_set: str):
    writer = ExcelWriter(TEST_FILEPATH)
    writer = set_field_value(writer, field, value_to_set)
    cell_value = str(writer.cell(0, field.value.cell_id).value)
    assert value_to_set in cell_value


def test_empty_class_set_to_not_involved():
    writer = ExcelWriter(TEST_FILEPATH)
    writer = set_field_value(writer, Field.CLASS, "")
    cell_value = str(writer.cell(0, Field.CLASS.value.cell_id).value)
    assert "Not Involved" in cell_value
