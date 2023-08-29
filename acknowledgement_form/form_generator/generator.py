from typing import List

from loguru import logger

from acknowledgement_form.form_generator.constants import (
    FIRST_CONTENT_DESCRIPTION_CELL,
    FIRST_CONTENT_TITLE_CELL,
    SIGNATURE_BLOCK_CELL_RANGE,
    TEMPLATE_FILEPATH,
    Content,
    Field,
)
from excel_writer.writer import ExcelWriter


def load_template(template_filepath: str = TEMPLATE_FILEPATH) -> ExcelWriter:
    return ExcelWriter(
        template_filepath, default_row_height=15.75, default_column_width=51.43
    )


def set_field_value(
    writer: ExcelWriter, field: Field, value_to_set: str
) -> ExcelWriter:
    if not value_to_set:
        value_to_set = "Not Involved" if field == Field.CLASS else "-"
    field_cell_id = field.value.cell_id
    field_template_str = field.value.template_str

    client_name_cell = writer.cell(0, field_cell_id)

    template_value = str(client_name_cell.value)

    new_value = template_value.replace(field_template_str, value_to_set)
    writer.cell(0, field_cell_id, new_value)

    return writer


def set_content(writer: ExcelWriter, contents: List[Content]) -> ExcelWriter:
    content_title_style = writer.cell_style(0, FIRST_CONTENT_TITLE_CELL)
    content_description_style = writer.cell_style(0, FIRST_CONTENT_DESCRIPTION_CELL)

    start_cell = writer.cell(0, FIRST_CONTENT_TITLE_CELL)

    current_row = start_cell.row
    column = start_cell.column

    current_signature_range = SIGNATURE_BLOCK_CELL_RANGE

    current_signature_range = writer.move_range(
        0, current_signature_range, rows_to_move=100
    )

    for title, descriptions in contents:
        writer.cell(
            0,
            cell_id=(current_row, column),
            set_value=title,
            set_style=content_title_style,
        )

        for description_line in descriptions:
            current_row += 1
            writer.cell(
                0,
                cell_id=(current_row, column),
                set_value=description_line,
                set_style=content_description_style,
            )

        current_row += 2

    rows_to_move = current_row - current_signature_range.start_row
    current_signature_range = writer.move_range(
        0, current_signature_range, rows_to_move=rows_to_move
    )

    return writer
