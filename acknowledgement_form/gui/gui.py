from typing import Dict, List, Optional, Tuple, Union

from appJar import gui
from loguru import logger

from acknowledgement_form.form_generator.constants import Content, Field
from acknowledgement_form.form_generator.email import ConfirmationEmailGenerator
from acknowledgement_form.form_generator.generator import (
    load_template,
    set_content,
    set_field_value,
)
from acknowledgement_form.form_generator.quotation_reader import QuotationReader
from excel_writer.writer import ExcelWriter


class AcknowledgementFormGeneratorGUI:
    _LABEL_ENTRIES: Dict[Union[str, Field], str] = {
        "pdf_filepath": "Quotation PDF Filepath:",
        Field.CLIENT_NAME: "Client Name:",
        Field.JOB_NUM: "Job No.:",
        Field.PO_NUM: "PO No.:",
        Field.QUOTATION_NUM: "Quotation No.:",
        Field.VESSEL: "Vessel:",
        Field.DRAWING_NUM: "Drawing No.:",
        Field.CLASS: "Class:",
        Field.DURATION: "Duration:",
        "output_filename": "Output File Name:",
    }

    _EMAIL_SUBJECT_TEXT_AREA_ID = "email_subject"
    _EMAIL_BODY_TEXT_AREA_ID = "email_body"
    _CONTENT_TEXT_AREA_ID = "contents"
    _TITLE_LINE = "Title: "
    _DESCRIPTION_START_LINE = "Description:\n"
    _DESCRIPTION_END_BLOCK = "---ENDBLOCK---"

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
        for entry_label in self._LABEL_ENTRIES.values():
            self.app.addLabelEntry(entry_label)
        self.app.addLabel("content_label", "Contents: ")
        self.app.setSticky("news")
        self.app.addScrolledTextArea(self._CONTENT_TEXT_AREA_ID)
        self.app.addLabel("email_subject_label", "Email Subject:")
        self.app.setSticky("esw")
        self.app.addTextArea(self._EMAIL_SUBJECT_TEXT_AREA_ID)
        self.app.addLabel("email_body_label", "Email Body")
        self.app.setSticky("news")
        self.app.addScrolledTextArea(self._EMAIL_BODY_TEXT_AREA_ID)

    def _setup_change_functions(self):
        for entry_label in self._LABEL_ENTRIES.values():
            self.app.setEntryChangeFunction(entry_label, self._update_fields)  # type: ignore
        self.app.setTextAreaChangeFunction(
            self._CONTENT_TEXT_AREA_ID, self._update_fields
        )

    def _setup_buttons(self):
        self.app.addButton("Select Quotation PDF", self._pdf_file_select)
        self.app.addButton("Save", self._save_file)
        self.app.addButton("Save As PDF", self._save_as_pdf)

    def _setup_icon(self):
        try:
            self.app.setImageLocation("acknowledgement_form/gui/resources/")
            self.app.setIcon("mencast_logo.gif")
        except Exception as exc:
            logger.warning(f"failed to set icon with error: {exc}")

    def _setup_size(self):
        self.app.setSize("1000x1000")

    def _pdf_file_select(self, button):
        if filepath := self.app.openBox(
            fileTypes=[("PDF Files", "*.pdf")], asFile=False
        ):
            self.app.setEntry(self._LABEL_ENTRIES["pdf_filepath"], filepath)
            self._populate_fields_from_quotation()

    def _populate_fields_from_quotation(self):
        quotation_filepath = self._get_entry("pdf_filepath")
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
            self.app.setEntry(self._LABEL_ENTRIES[field], value)

    def _set_contents_into_text_area(self, contents: List[Content]) -> None:
        content_text = ""
        for title, descriptions in contents:
            content_text = self._update_content_text(content_text, title, descriptions)
            content_text += "\n"
        self.app.setTextArea(self._CONTENT_TEXT_AREA_ID, content_text)

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
        content_text += f"{self._TITLE_LINE}{title} \n"
        content_text += self._DESCRIPTION_START_LINE
        return content_text

    def _update_content_text_with_descriptions(
        self, content_text: str, descriptions: List[str]
    ) -> str:
        for description_line in descriptions:
            content_text += f"{description_line} \n"
        content_text += f"\n{self._DESCRIPTION_END_BLOCK}\n"
        return content_text

    def _update_fields(self):
        contents = self._get_content_from_text_area()
        self._set_output_filename()
        self._update_email_text_area(contents)

    def _set_output_filename(self):
        output_filename = self._generate_output_filename_from_fields()
        self.app.setEntry(self._LABEL_ENTRIES["output_filename"], output_filename)

    def _generate_output_filename_from_fields(self):
        client_name = self._get_entry(Field.CLIENT_NAME)
        job_number = self._get_entry(Field.JOB_NUM)
        quotation_number = self._get_entry(Field.QUOTATION_NUM)
        return f"ACK-JN{job_number}-{quotation_number}-{client_name}".replace(" ", "_")

    def _save_file(self, button):
        if not self._check_entries_not_empty():
            return
        self._generate_acknowledgement()
        excel_filepath = self._get_excel_filepath()
        if excel_filepath is not None:
            self.writer.save_workbook("", filename=excel_filepath)

    def _save_as_pdf(self, button):
        if not self._check_entries_not_empty():
            return
        self._generate_acknowledgement()
        pdf_filepath = self._get_pdf_filepath()
        if pdf_filepath is not None:
            self.writer.export_as_pdf("", filename=pdf_filepath)

    def _check_entries_not_empty(self) -> bool:
        if missing_entries := [
            entry_label
            for entry_label in self._LABEL_ENTRIES.values()
            if not self._get_entry(entry_label)
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

    def _generate_acknowledgement(self) -> bool:
        self._load_writer()
        self._update_writer_with_gui_contents()
        return True

    def _update_writer_with_gui_contents(self):
        for field in Field:
            field_value = self._get_entry(field)
            self._set_field_value(field, field_value)
        self._set_contents()

    def _get_content_from_text_area(self) -> List[Content]:
        contents = []
        contents = str(self.app.getTextArea(self._CONTENT_TEXT_AREA_ID))
        content_lines = contents.split("\n")
        titles = [
            line[len(self._TITLE_LINE) :]
            for line in content_lines
            if line.startswith(self._TITLE_LINE)
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
            if self._is_description_start_line(line):
                continue
            if self._is_end_of_description_block(line):
                descriptions.append(
                    self._get_description_lines(
                        description_lines, description_line_indexes
                    )
                )
                description_line_indexes.clear()
                continue
            description_line_indexes.append(line_index)

        return descriptions

    def _is_end_of_description_block(self, line: str) -> bool:
        return line == self._DESCRIPTION_END_BLOCK

    def _extract_description_lines(self, content_lines: List[str]) -> List[str]:
        return [line for line in content_lines if not line.startswith(self._TITLE_LINE)]

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

    def _get_excel_filepath(self) -> Optional[str]:
        return self._get_filepath(("Excel", "*.xlsx"))

    def _get_pdf_filepath(self) -> Optional[str]:
        return self._get_filepath(("PDF", "*.pdf"))

    def _get_filepath(self, filetype: Tuple[str, str]) -> Optional[str]:
        output_filename = self._read_output_entry()
        if filepath := self._open_savebox(
            output_filename=output_filename, filetype=filetype
        ):
            return filepath
        return None

    def _open_savebox(
        self,
        output_filename: Optional[str] = None,
        directory: Optional[str] = None,
        filetype: Optional[Tuple[str, str]] = None,
    ) -> str:
        return self.app.saveBox(
            fileName=output_filename, dirName=directory, fileTypes=[filetype]
        )

    def _read_output_entry(self) -> str:
        return self._get_entry("output_filename")

    def go(self):
        self.app.go()

    def _get_entry(self, entry_id: Union[Field, str]) -> str:
        if entry_id in self._LABEL_ENTRIES:
            return str(self.app.getEntry(self._LABEL_ENTRIES[entry_id]))
        return str(self.app.getEntry(entry_id))

    def _update_email_text_area(self, contents: List[Content]) -> None:
        self._reset_email_text_area()

        email_generator = self._instantiate_email_generator(contents)

        subject = email_generator.create_email_subject()
        body = email_generator.create_email_body()

        self.app.setTextArea(self._EMAIL_SUBJECT_TEXT_AREA_ID, subject)
        self.app.setTextArea(self._EMAIL_BODY_TEXT_AREA_ID, body)

    def _instantiate_email_generator(
        self, contents: List[Content]
    ) -> ConfirmationEmailGenerator:
        entry_value_map = {field: self._get_entry(field) for field in Field}
        return ConfirmationEmailGenerator(
            client_name=entry_value_map[Field.CLIENT_NAME],
            vessel_details=entry_value_map[Field.VESSEL],
            quotation_number=entry_value_map[Field.QUOTATION_NUM],
            job_number=entry_value_map[Field.JOB_NUM],
            contents=contents,
            duration=entry_value_map[Field.DURATION],
            vessel_class=entry_value_map[Field.CLASS],
            po_number=entry_value_map[Field.PO_NUM],
            drawing_number=entry_value_map[Field.DRAWING_NUM],
        )

    def _reset_email_text_area(self) -> None:
        self._reset_text_area(self._EMAIL_SUBJECT_TEXT_AREA_ID)
        self._reset_text_area(self._EMAIL_BODY_TEXT_AREA_ID)

    def _reset_text_area(self, text_area_id: str) -> None:
        self.app.clearTextArea(text_area_id)
