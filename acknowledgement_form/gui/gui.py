from typing import Dict, List, Optional, Tuple, Union

from appJar import gui
from loguru import logger

from acknowledgement_form.form_generator.constants import Content, Field
from acknowledgement_form.form_generator.generator import (
    load_template,
    set_content,
    set_field_value,
)
from acknowledgement_form.form_generator.quotation_reader import QuotationReader
from excel_writer.writer import ExcelWriter


class AcknowledgementFormGeneratorGUI:
    LABEL_ENTRIES: List[Tuple[Union[str, Field], str]] = [
        ("pdf_filepath", "Quotation PDF Filepath:"),
        (Field.CLIENT_NAME, "Client Name:"),
        (Field.JOB_NUM, "Job No.:"),
        (Field.PO_NUM, "PO No.:"),
        (Field.QUOTATION_NUM, "Quotation No.:"),
        (Field.VESSEL, "Vessel:"),
        (Field.DRAWING_NUM, "Drawing No.:"),
        (Field.CLASS, "Class:"),
        (Field.DURATION, "Duration:"),
        ("output_filename", "Output File Name:"),
    ]

    CONTENT_TEXT_AREA_ID = "contents"
    TITLE_LINE = "Title: "
    DESCRIPTION_START_LINE = "Description:\n"

    def __init__(self):
        self.app = gui("Acknowledgement Form Generator", useTtk=True)
        self.writer: ExcelWriter
        self._setup_gui()

    def _setup_gui(self):
        self._setup_labels()
        self._setup_entries()
        self._setup_buttons()
        self._setup_icon()
        self._setup_size()
        self._setup_change_functions()

    def _setup_labels(self):
        self.app.addLabel(
            "instructions", "Creates Job Acknowledgement Excel from Quotation PDF"
        )

    def _setup_entries(self):
        for entry_id, entry_label in self.LABEL_ENTRIES:
            self.app.addLabelEntry(entry_id, label=entry_label)
        self.app.addLabel("content_label", "Contents: ")
        self.app.addScrolledTextArea(self.CONTENT_TEXT_AREA_ID)

    def _setup_change_functions(self):
        for entry_id, _ in self.LABEL_ENTRIES:
            self.app.setEntryChangeFunction(entry_id, self._update_fields)  # type: ignore

    def _setup_buttons(self):
        self.app.addButton("Select Quotation PDF", self._pdf_file_select)
        self.app.addButton("Save", self._save_file)

    def _setup_icon(self):
        try:
            self.app.setImageLocation("acknowledgement_form/gui/resources/")
            self.app.setIcon("mencast_logo.gif")
        except Exception as exc:
            logger.warning(f"failed to set icon with error: {exc}")

    def _setup_size(self):
        self.app.setSize("500x700")

    def _pdf_file_select(self, button):
        if filepath := self.app.openBox(
            fileTypes=[("PDF Files", "*.pdf")], asFile=False
        ):
            self.app.setEntry("pdf_filepath", filepath)
            self._populate_fields_from_quotation()

    def _populate_fields_from_quotation(self):
        quotation_filepath = str(self.app.getEntry("pdf_filepath"))
        try:
            self._update_ui_with_quotation_data(quotation_filepath)
        except Exception as exc:
            logger.warning(f"failed to get values from quotation pdf: {exc}")

    def _update_ui_with_quotation_data(self, quotation_filepath: str) -> None:
        reader = QuotationReader(quotation_filepath)
        fields = reader.get_fields()
        self._set_fields_into_label_entries(fields)
        contents = reader.get_content()
        self._set_contents_into_text_area(contents)

    def _set_fields_into_label_entries(self, fields: Dict[Field, str]) -> None:
        for field, value in fields.items():
            self.app.setEntry(field, value)

    def _set_contents_into_text_area(self, contents: List[Content]) -> None:
        content_text = ""
        for title, descriptions in contents:
            content_text = self._update_content_text(content_text, title, descriptions)
            content_text += "\n"
        self.app.setTextArea(self.CONTENT_TEXT_AREA_ID, content_text)

    def _update_content_text(
        self, content_text: str, title: str, descriptions: List[str]
    ) -> str:
        content_text = self._update_content_text_with_title_line(content_text, title)
        content_text = self._update_content_text_with_descriptions(
            content_text, descriptions
        )
        return content_text

    def _update_content_text_with_title_line(
        self, content_text: str, title: str
    ) -> str:
        content_text += f"{self.TITLE_LINE}{title} \n"
        content_text += self.DESCRIPTION_START_LINE
        return content_text

    def _update_content_text_with_descriptions(
        self, content_text: str, descriptions: List[str]
    ) -> str:
        for description_line in descriptions:
            content_text += f"{description_line} \n"
        return content_text

    def _update_fields(self):
        self._set_output_filename()

    def _set_output_filename(self):
        output_filename = self._generate_output_filename_from_fields()
        self.app.setEntry("output_filename", output_filename)

    def _generate_output_filename_from_fields(self):
        client_name = str(self.app.getEntry(Field.CLIENT_NAME))
        job_number = str(self.app.getEntry(Field.JOB_NUM))
        quotation_number = str(self.app.getEntry(Field.QUOTATION_NUM))
        return f"ACK - JN{job_number} - {quotation_number} - {client_name}"

    def _save_file(self, button):
        if not self._check_entries_not_empty():
            return
        self._load_writer()
        self._generate_job_ack()
        full_filepath = self._get_full_output_filepath()
        if full_filepath is not None:
            self.writer.save_workbook("", filename=full_filepath)
            logger.info(f"saved workbook to {full_filepath}")

    def _check_entries_not_empty(self) -> bool:
        if missing_entries := [
            entry_label
            for entry_id, entry_label in self.LABEL_ENTRIES
            if not str(self.app.getEntry(entry_id))
        ]:
            self._warning_window_for_missing_entries(missing_entries)
            return False
        return True

    def _warning_window_for_missing_entries(self, missing_entries: List[str]):
        missing_entries_str = "\n".join(missing_entries)
        self.app.warningBox(
            title="Missing Data",
            message=f"Entries are missing: \n\n{missing_entries_str}",
        )

    def _load_writer(self):
        self.writer = load_template()

    def _generate_job_ack(self):
        for field in Field:
            field_value = str(self.app.getEntry(field))
            self._set_field_value(field, field_value)
        self._set_contents()

    def _get_content_from_text_area(self) -> List[Content]:
        contents = []
        contents = str(self.app.getTextArea(self.CONTENT_TEXT_AREA_ID))
        content_lines = contents.split("\n")
        titles = [
            line[len(self.TITLE_LINE) :]
            for line in content_lines
            if line.startswith(self.TITLE_LINE)
        ]
        descriptions = self._parse_descriptions_from_lines(content_lines)
        return [
            Content(title, description)
            for title, description in zip(titles, descriptions)
        ]

    def _parse_descriptions_from_lines(
        self, content_lines: List[str]
    ) -> List[List[str]]:
        description_lines = self._extract_description_lines(content_lines)

        descriptions: List[List[str]] = []
        description_line_indexes: List[int] = []

        for line_index, line in enumerate(description_lines):
            if line_index == 0:
                continue
            if self._is_description_start_line(line):
                descriptions.append(
                    self._get_description_lines(
                        description_lines, description_line_indexes
                    )
                )
                description_line_indexes.clear()
            else:
                description_line_indexes.append(line_index)
        return descriptions

    def _extract_description_lines(self, content_lines: List[str]) -> List[str]:
        return [line for line in content_lines if not line.startswith(self.TITLE_LINE)]

    def _is_description_start_line(self, line: str) -> bool:
        return line.startswith("Description:")

    def _get_description_lines(
        self, content_lines: List[str], description_line_indexes: List[int]
    ) -> List[str]:
        return [
            content_lines[index]
            for index in description_line_indexes
            if content_lines[index].strip()
        ]

    def _set_field_value(self, field: Field, value_to_set: str):
        self.writer = set_field_value(self.writer, field, value_to_set)

    def _set_contents(self):
        contents = self._get_content_from_text_area()
        self.writer = set_content(self.writer, contents)

    def _get_full_output_filepath(self) -> Optional[str]:
        output_filename = str(self.app.getEntry("output_filename"))
        if filepath := self.app.saveBox(
            fileName=output_filename,
            dirName=".",
            fileTypes=[("Excel Workbook", "*.xlsx")],
        ):
            return filepath
        return None

    def go(self):
        self.app.go()
