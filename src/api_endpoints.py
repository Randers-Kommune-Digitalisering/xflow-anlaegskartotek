"""
Receive XFlox asset and leasing reports.
"""
import logging
import os

from flask import Blueprint, Response, current_app, request

from utils.document_handler import download_contracts
from utils.email_handler import send_report_email
from utils.excel_handler import create_xflow_anlaeg_excel

api = Blueprint("xflow_anlaeg", __name__, url_prefix="/api")
logger = logging.getLogger(__name__)


@api.route("/xflow-anlaeg", methods=["GET"])
def get_xflow_anlaeg_data() -> Response:
    """
    Receive an asset report submitted through XFLow.
    """
    # Allow mock payload during local development
    if current_app.debug and os.getenv("XFLOW_ANLAEG_MOCK") == "1":
        payload = {
            "regnskabsaar": "2026",
            "institution/afdeling": "Testinstitution",
            "kontaktperson": "Testperson",
            "telefon": "12345678",
            "dato": "11-12-2026",
            "grunde_og_bygninger": {
                "anskaffet_fast_ejendom": "Nej",
                "nyanskaffet_ejendom": [],
                "afhaendet_fast_ejendom": "Nej",
                "nyafhaendet_ejendom": [],
                "foretaget_forbedringer": "Nej",
                "forbedringer": [],
            },
            "tekniske_anlaeg": {
                "anskaffet_anlaeg_100000": "Nej",
                "nyanskaffet_anlaeg_100000": [],
                "anskaffet_anlaeg": "Ja",
                "nyanskaffet_anlaeg": [
                    {
                        "nyanskaffet_a_beskrivelse": "Udskiftning af cirkulationspumpe i idrætshal",
                        "nyanskaffet_a_bilagsnr": "1000000000",
                        "nyanskaffet_a_pris": "32500",
                        "nyanskaffet_a_aargang": "2026",
                    },
                ],
            },
            "inventar": {
                "anskaffet_inventar_100000": "Nej",
                "nyanskaffet_inventar_100000": [],
                "anskaffet_inventar": "Ja",
                "nyanskaffet_inventar": [
                    {
                        "nyanskaffet_i_beskrivelse": "nye omklædnngsskabe",
                        "nyanskaffet_i_bilagsnr": "2000000000",
                        "nyanskaffet_i_pris": "18500",
                        "nyanskaffet_i_aargang": "2026",
                    },
                ],
            },
            "leasingaftaler": {
                "indgaaet_leasingaftaler": "Nej",
                "indgaaede_leasingaftaler": [],
                "aftalekontrakter": [],
                "opsagt_leasingaftaler": "Nej",
                "opsagte_leasingaftaler": [],
                "andre_bemaerkninger": "Der er ikke indgået eller opsagt leasingaftaler i regnskabsåret. Der er derfor ingen leasingkontrakt at vedhæfte.",
            },
        }

        logger.info("Using XFlow asset mock payload")

    else:
        payload = request.get_json(silent=True)

        if not isinstance(payload, dict):
            return Response(
                "Invalid payload: expected a JSON object", status=400,
            )

        logger.debug("XFlow payload: %s", payload)

    # Validate required top-level fields
    required_fields = [
        "regnskabsaar",
        "institution/afdeling",
        "kontaktperson",
        "telefon",
        "dato",
        "grunde_og_bygninger",
        "tekniske_anlaeg",
        "inventar",
        "leasingaftaler",
    ]

    for field in required_fields:
        if field not in payload:
            logger.error("Missing required field: %s", field)
            return Response(
                f"Invalid payload: missing required field '{field}'", status=400,
            )

    # Validate sections
    sections = [
        "grunde_og_bygninger",
        "tekniske_anlaeg",
        "inventar",
        "leasingaftaler",
    ]

    for section in sections:
        if not isinstance(payload[section], dict):
            return Response(
                f"Invalid payload: '{section}' must be a JSON object", status=400,
            )

    # Validate repeating fields
    array_fields = {
        "grunde_og_bygninger": [
            "nyanskaffet_ejendom",
            "nyafhaendet_ejendom",
            "forbedringer",
        ],
        "tekniske_anlaeg": [
            "nyanskaffet_anlaeg_100000",
            "nyanskaffet_anlaeg",
        ],
        "inventar": [
            "nyanskaffet_inventar_100000",
            "nyanskaffet_inventar",
        ],
        "leasingaftaler": [
            "indgaaede_leasingaftaler",
            "aftalekontrakter",
            "opsagte_leasingaftaler",
        ],
    }

    for section, fields in array_fields.items():
        for field in fields:
            value = payload[section].get(field)

            if not isinstance(value, list):
                return Response(
                    f"Invalid payload: '{section}.{field}' must me an array.", status=400,
                )

    # Extract information
    regnskabsaar = payload["regnskabsaar"]
    institution = payload["institution/afdeling"]

    leasingaftaler = payload["leasingaftaler"]
    contracts = leasingaftaler["aftalekontrakter"]

    logger.info(
        "XFlow asset report received: year=%s, institution=%s",
        regnskabsaar,
        institution,
    )

    logger.info(
        "Number of attached leasing contracts: %s",
        len(contracts),
    )

    try:
        excel_bytes = create_xflow_anlaeg_excel(payload)
    except (KeyError, TypeError, ValueError):
        logger.exception("Failed to generate XFlow Excel report")
        return Response(
            "Excel report could not be generated", status=500,
        )

    filename = (
        f"Anlaegsindberetning_"
        f"{payload['institution/afdeling']}_"
        f"{payload['regnskabsaar']}.xlsx"
    )

    logger.info("XFlow Excel report generated: filename=%s", filename)

    contracts_downloaded = download_contracts(payload) if contracts else []

    email_attachments = [(filename, excel_bytes)]
    email_attachments.extend(
        (contract["filename"], contract["content"])
        for contract in contracts_downloaded
    )

    return Response("Report generated; email disabled", status=200)

    # send_report_email(email_attachments)
    # return Response("Report sent", status=200)
