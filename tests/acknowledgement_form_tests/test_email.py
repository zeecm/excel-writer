from typing import List, Optional

import pytest

from acknowledgement_form.form_generator.constants import Content
from acknowledgement_form.form_generator.email import ConfirmationEmailGenerator


@pytest.mark.parametrize(
    "client_name, vessel_details, quotation_number, job_number, po_number,"
    " expected_subject",
    [
        (
            "ABC PTE LTD",
            "15M Boat",
            "MMSQ23-123",
            "2308001",
            None,
            "Confirmation: ABC PTE LTD - 15M Boat - MMSQ23-123 - JOB NO 2308001",
        ),
        (
            "XYZ PTE LTD",
            "40M Ship",
            "MMSQ23-456",
            "2308099",
            "12345",
            (
                "Confirmation: XYZ PTE LTD - 40M Ship - MMSQ23-456 - PO 12345 - JOB NO"
                " 2308099"
            ),
        ),
    ],
)
def test_email_subject(
    client_name: str,
    vessel_details: str,
    quotation_number: str,
    job_number: str,
    po_number: Optional[str],
    expected_subject: str,
):
    email_generator = ConfirmationEmailGenerator(
        client_name=client_name,
        vessel_details=vessel_details,
        quotation_number=quotation_number,
        job_number=job_number,
        po_number=po_number,
    )
    subject = email_generator.create_email_subject()
    assert subject == expected_subject


@pytest.mark.parametrize(
    "client_name, vessel_details, quotation_number, job_number, contents, duration,"
    " vessel_class, po_number, expected_body",
    [
        (
            "ABC PTE LTD",
            "15M Boat",
            "MMSQ23-123",
            "2308001",
            [
                Content(
                    "title1",
                    [
                        "desc1",
                        "desc2",
                    ],
                ),
                Content(
                    "title2",
                    [
                        "desc1",
                        "desc2",
                    ],
                ),
            ],
            "2 months",
            "BV",
            None,
            (
                "Hi All,\n\nPlease take note of Confirmation: ABC PTE LTD - 15M Boat -"
                " MMSQ23-123 - JOB NO 2308001.\n\nJob No.:"
                " 2308001\n\nContent:\n\ntitle1\ndesc1\ndesc2\n\ntitle2\ndesc1\ndesc2\n\nDuration:"
                " 2 months\nClass: BV\n\nFor work detail please refer to attached"
                " quotation.\n\n"
            ),
        ),
        (
            "XYZ PTE LTD",
            "40M Ship",
            "MMSQ23-456",
            "2308099",
            [
                Content(
                    "title1",
                    [
                        "desc1",
                    ],
                ),
            ],
            "2 days",
            "ABS",
            "123456",
            (
                "Hi All,\n\nPlease take note of Confirmation: XYZ PTE LTD - 40M Ship -"
                " MMSQ23-456 - PO 123456 - JOB NO 2308099.\n\nJob No.: 2308099\nPO No.:"
                " 123456\n\nContent:\n\ntitle1\ndesc1\n\nDuration: 2 days\nClass:"
                " ABS\n\nFor work detail please refer to attached quotation and PO.\n\n"
            ),
        ),
    ],
)
def test_email_body(
    client_name: str,
    vessel_details: str,
    quotation_number: str,
    job_number: str,
    contents: List[Content],
    duration: str,
    vessel_class: str,
    po_number: Optional[str],
    expected_body: str,
):
    email_generator = ConfirmationEmailGenerator(
        client_name=client_name,
        vessel_details=vessel_details,
        quotation_number=quotation_number,
        job_number=job_number,
        contents=contents,
        duration=duration,
        vessel_class=vessel_class,
        po_number=po_number,
    )
    email_body = email_generator.create_email_body()
    assert email_body == expected_body
