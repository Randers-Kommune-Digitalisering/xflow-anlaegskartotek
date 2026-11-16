import os

from rkdigi.email_handling import EmailSender


def send_report_email(attachments: list[tuple[str, bytes]]) -> None:
    """
    Send the asset report and any leasing contracts

    :param attachments: List of tuples containing the filename and content of each attachment
    """
    EmailSender(
        smtp_server=os.environ["SMTP_SERVER"],
        smtp_port=int(os.environ["SMTP_PORT"]),
        sender_email=os.environ["SMTP_SENDER"],
        sender_password=os.environ["SMTP_PASSWORD"],
    ).send_email(
        recipients=os.environ["REPORT_EMAIL_RECIPIENT"],
        subject="Anlægsindberetning",
        body="Anlægsindberetning er vedhæftet.",
        attachments=attachments,
    )
