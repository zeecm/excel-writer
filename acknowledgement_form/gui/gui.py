from appJar import gui
from loguru import logger
from pandas import ExcelWriter

from acknowledgement_form.form_generator.constants import Field
from acknowledgement_form.form_generator.generator import load_template, set_field_value
from acknowledgement_form.form_generator.quotation_reader import (
    get_fields_from_quotation_pdf,
)


class AcknowledgementFormGeneratorGUI:
    def __init__(self):
        self.app = gui("Acknowledgement Form Generator", useTtk=True)
        self.writer = ExcelWriter
        self._setup_gui()

    def _setup_gui(self):
        self._setup_labels()
        self._setup_entries()
        self._setup_buttons()
        self._setup_icon()
        self._setup_size()
        self._setup_theme()
        self._setup_change_functions()

    def _setup_labels(self):
        self.app.addLabel(
            "instructions", "Creates Job Acknowledgement Excel from Quotation PDF"
        )

    def _setup_entries(self):
        self.app.addLabelEntry("pdf_filepath", label="Quotation PDF Filepath:")
        self.app.addLabelEntry(Field.CLIENT_NAME, label="Client Name:")
        self.app.addLabelEntry(Field.JOB_NUM, label="Job No.:")
        self.app.addLabelEntry(Field.PO_NUM, label="PO No.:")
        self.app.addLabelEntry(Field.QUOTATION_NUM, label="Quotation No.:")
        self.app.addLabelEntry(Field.VESSEL, label="Vessel:")
        self.app.addLabelEntry(Field.DRAWING_NUM, label="Drawing No.:")
        self.app.addLabelEntry(Field.CLASS, label="Class:")
        self.app.addLabelEntry(Field.DURATION, label="Duration:")
        self.app.addLabelEntry("output_filename", label="Output File Name:")

    def _setup_change_functions(self):
        labels = ["pdf_filepath", "output_filename"] + list(Field)
        for label in labels:
            self.app.setEntryChangeFunction(label, self._update_fields)  # type: ignore

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

    def _setup_theme(self):
        self.app.setTtkTheme("clam")

    def _pdf_file_select(self, button):
        if filepath := self.app.openBox(
            fileTypes=[("PDF Files", "*.pdf")], asFile=False
        ):
            self.app.setEntry("pdf_filepath", filepath)
            self._populate_fields_from_quotation()

    def _populate_fields_from_quotation(self):
        quotation_filepath = self.app.getEntry("pdf_filepath")
        try:
            fields = get_fields_from_quotation_pdf(quotation_filepath)
            for field, value in fields.items():
                self.app.setEntry(field, value)
        except Exception as exc:
            logger.warning(f"failed to get values from quotation pdf: {exc}")

    def _generate_output_filename_from_fields(self):
        client_name = str(self.app.getEntry(Field.CLIENT_NAME))
        job_number = str(self.app.getEntry(Field.JOB_NUM))
        quotation_number = str(self.app.getEntry(Field.QUOTATION_NUM))
        return f"ACK - JN{job_number} - {quotation_number} - {client_name}"

    def _set_output_filename(self):
        output_filename = self._generate_output_filename_from_fields()
        self.app.setEntry("output_filename", output_filename)

    def _update_fields(self):
        self._set_output_filename()

    def _generate_job_ack(self):
        for field in Field:
            field_value = str(self.app.getEntry(field))
            self._set_field_value(field, field_value)

    def _set_field_value(self, field: str, value_to_set: str):
        self.writer = set_field_value(self.writer, field, value_to_set)

    def _save_file(self, button):
        self._load_writer()
        self._generate_job_ack()
        full_filepath = self._get_full_output_filepath()
        self.writer.save_workbook("", filename=full_filepath)
        logger.info(f"saved workbook to {full_filepath}")

    def _get_full_output_filepath(self):
        output_filename = str(self.app.getEntry("output_filename"))
        if filepath := self.app.saveBox(
            fileName=output_filename,
            dirName=".",
            fileTypes=[("Excel Workbook", "*.xlsx")],
        ):
            return filepath
        return f"./{output_filename}"

    def _load_writer(self):
        self.writer = load_template()

    def go(self):
        self.app.go()
