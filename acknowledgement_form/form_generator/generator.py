from acknowledgement_form.form_generator.constants import TEMPLATE_FILEPATH, Field
from acknowledgement_form.form_generator.quotation_reader import (
    get_fields_from_quotation_pdf,
)
from excel_writer.writer import ExcelWriter


def load_template(template_filepath: str = TEMPLATE_FILEPATH) -> ExcelWriter:
    return ExcelWriter(template_filepath)


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


def create_job_ack_from_pdf(
    pdf_filepath: str,
    filename: str,
    filepath: str = ".",
    template_filepath: str = TEMPLATE_FILEPATH,
) -> None:
    fields = get_fields_from_quotation_pdf(pdf_filepath)
    writer = load_template(template_filepath)
    for field, value in fields.items():
        writer = set_field_value(writer, field, value)
    writer.save_workbook(filepath, filename)


if __name__ == "__main__":
    test_filename = "test.xlsx"
    test_pdf_filepath = (
        "tests/acknowledgement_form_tests/test_files/sample_quo_with_version.pdf"
    )
    create_job_ack_from_pdf(pdf_filepath=test_pdf_filepath, filename=test_filename)
