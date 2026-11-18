import os

from rkdigi.email_handling import EmailSender
from utils.config import RECIPIENTS, SMTP_SERVER, SMTP_SENDER


def send_report_email(attachments: list[tuple[str, bytes]]) -> None:
    """
    Send the asset report and any leasing contracts

    :param attachments: List of tuples containing the filename and content of each attachment
    """
    EmailSender(
        smtp_server=SMTP_SERVER,
    ).send_email(
        recipients=[RECIPIENTS],
        sender=SMTP_SENDER,
        subject="Anlægsindberetning",
        body="Anlægsindberetning er vedhæftet.",
        attachments=attachments,
    )
