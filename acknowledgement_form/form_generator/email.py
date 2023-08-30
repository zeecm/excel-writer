from typing import Dict, List, Optional

from acknowledgement_form.form_generator.constants import Content


class ConfirmationEmailGenerator:
    def __init__(
        self,
        client_name: str = "",
        vessel_details: str = "",
        quotation_number: str = "",
        job_number: str = "",
        contents: Optional[List[Content]] = None,
        duration: str = "",
        vessel_class: str = "Not Involved",
        po_number: Optional[str] = None,
    ):
        contents = contents or []
        self._client_name = client_name
        self._vessel_details = vessel_details
        self._quotation_number = quotation_number
        self._job_number = job_number
        self._contents = contents
        self._vessel_class = vessel_class
        self._po_number = po_number
        self._duration = duration

    def create_email_subject(self) -> str:
        subject = "Confirmation: "
        info_string_map = self._create_info_string_map()
        for info, info_str in info_string_map.items():
            if info is not None:
                subject += info_str
        return subject

    def _create_info_string_map(self) -> Dict[Optional[str], str]:
        return {
            self._client_name: f"{self._client_name} - ",
            self._vessel_details: f"{self._vessel_details} - ",
            self._quotation_number: f"{self._quotation_number} - ",
            self._po_number: f"PO {self._po_number} - ",
            self._job_number: f"JOB NO {self._job_number}",
        }

    def create_email_body(self) -> str:
        subject = self.create_email_subject()
        content_lines = self.create_lines_from_contents()
        final_line = self._create_final_line()

        return self._generate_email_body_lines(
            subject=subject,
            content_lines=content_lines,
            final_line=final_line,
        )

    def create_lines_from_contents(self) -> str:
        lines = []
        for title, descriptions in self._contents:
            lines.extend((title, "\n"))
            for description in descriptions:
                lines.extend((description, "\n"))
            lines.append("\n")
        return "".join(lines[:-2])

    def _create_final_line(self) -> str:
        final_line = "For work detail please refer to attached quotation"
        if self._po_number is not None:
            final_line += " and PO"
        return f"{final_line}."

    def _generate_email_body_lines(
        self,
        subject: str,
        content_lines: str,
        final_line: str,
    ) -> str:
        email_body_lines = [
            "Hi All,",
            "\n",
            "\n",
            f"Please take note of {subject}.",
            "\n",
            "\n",
            f"Job No.: {self._job_number}",
            f"\nPO No.: {self._po_number}" if self._po_number is not None else "",
            "\n",
            "\n",
            "Content:",
            "\n",
            "\n",
            content_lines,
            "\n",
            "\n",
            f"Duration: {self._duration}",
            "\n",
            "\n",
            f"Class: {self._vessel_class}",
            "\n",
            "\n",
            final_line,
            "\n",
            "\n",
        ]
        return "".join(email_body_lines)
