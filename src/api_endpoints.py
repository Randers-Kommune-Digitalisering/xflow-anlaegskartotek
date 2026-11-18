"""
Receive XFlox asset and leasing reports.
"""
import logging
import os

from flask import Blueprint, Response, current_app, request

from utils.document_handler import download_contracts
from utils.email_handler import send_report_email
from utils.excel_handler import create_xflow_anlaeg_excel
from flask import Blueprint, Response, current_app, request, send_file

api = Blueprint("xflow_anlaeg", __name__, url_prefix="/api")
logger = logging.getLogger(__name__)


@api.route("/xflow-anlaeg", methods=["POST"])
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
                "anskaffet_fast_ejendom": "Ja",
                "nyanskaffet_ejendom": [
                    {
                        "anskaffet_e_beskrivelse": "Erhvervelse af naboejendom til udvidelse af institutionen",
                        "anskaffet_e_bilagsnr": "1000000001",
                        "anskaffet_e_dato": "15-03-2026",
                        "anskaffet_e_pris": "2500000",
                    },
                ],
                "afhaendet_fast_ejendom": "Ja",
                "nyafhaendet_ejendom": [
                    {
                        "afhaendet_e_beskrivelse": "Salg af tidligere administrationsbygning",
                        "afhaendet_e_bilagsnr": "1000000002",
                        "afhaendet_e_dato": "20-08-2026",
                        "afhaendet_e_pris": "1800000",
                    },
                ],
                "foretaget_forbedringer": "Ja",
                "forbedringer": [
                    {
                        "forbedring_beskrivelse": "Gennemgribende renovering af institutionens tag og facade",
                        "forbedring_bilagsnr": "1000000003",
                        "forbedring_dato": "01-06-2026",
                        "forbedring_pris": "450000",
                    },
                ],
            },

            "tekniske_anlaeg": {
                "anskaffet_anlaeg": "Ja",
                "nyanskaffet_anlaeg": [
                    {
                        "anskaffet_a_beskrivelse": "Udskiftning af cirkulationspumpe i idrætshal",
                        "anskaffet_a_bilagsnr": "2000000001",
                        "anskaffet_a_pris": "32500",
                        "anskaffet_a_aargang": "2026",
                    },
                    {
                        "anskaffet_a_beskrivelse": "Ny ventilationsmaskine til skolebygning",
                        "anskaffet_a_bilagsnr": "2000000002",
                        "anskaffet_a_pris": "175000",
                        "anskaffet_a_aargang": "2026",
                    },
                ],
                "afhaendet_anlaeg": "Ja",
                "nyafhaendet_anlaeg": [
                    {
                        "afhaendet_a_beskrivelse": "Afhændelse af ældre traktor",
                        "afhaendet_a_bilagsnr": "2000000003",
                        "afhaendet_a_pris": "45000",
                        "afhaendet_a_aargang": "2026",
                    },
                ]
            },

            "inventar": {
                "anskaffet_inventar": "Ja",
                "nyanskaffet_inventar": [
                    {
                        "anskaffet_i_beskrivelse": "Nye omklædningsskabe",
                        "anskaffet_i_bilagsnr": "3000000001",
                        "anskaffet_i_pris": "18500",
                        "anskaffet_i_aargang": "2026",
                    },
                    {
                        "anskaffet_i_beskrivelse": "Nye bærbare computere",
                        "anskaffet_i_bilagsnr": "3000000002",
                        "anskaffet_i_pris": "125000",
                        "anskaffet_i_aargang": "2026",
                    },
                ],
                "afhaendet_inventar": "Ja",
                "nyafhaendet_inventar": [
                    {
                        "afhaendet_i_beskrivelse": "Afhændelse af ældre computere",
                        "afhaendet_i_bilagsnr": "3000000003",
                        "afhaendet_i_pris": "5000",
                        "afhaendet_i_aargang": "2026",
                    },
                ],
            },

            "leasingaftaler": {
                "indgaaet_leasingaftaler": "Ja",
                "indgaaede_leasingaftaler": [
                    {
                        "indgaaet_beskrivelse": "Leasing af ny kopimaskine",
                        "indgaaet_bemaerkninger": "Leasingaftalen løber over 48 måneder.",
                    },
                ],
                "aftalekontrakter": [],
                "opsagt_leasingaftaler": "Ja",
                "opsagte_leasingaftaler": [
                    {
                        "opsagt_beskrivelse": "Opsigelse af leasingaftale på tidligere printere",
                        "opsagt_bemaerkninger": "Aftalen ophørte pr. 30-09-2026.",
                    },
                ],
                "andre_bemaerkninger": "Der er både indgået en ny leasingaftale og opsagt en eksisterende leasingaftale i regnskabsåret.",
            },
        }

        # Only for test purpose on the platform
        logger.debug("Using XFlow asset mock payload")

    else:
        logger.info("Using XFlow asset payload from xflow")
        payload = request.get_json(silent=True)

        if not isinstance(payload, dict):
            return Response(
                "Invalid payload: expected a JSON object", status=400,
            )

        logger.info("XFlow payload: %s", payload)

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
            "nyanskaffet_anlaeg",
            "nyafhaendet_anlaeg",
        ],
        "inventar": [
            "nyanskaffet_inventar",
            "nyafhaendet_inventar",
        ],
        "leasingaftaler": [
            "indgaaede_leasingaftaler",
            # "aftalekontrakter",
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

    # til lokal
    # return send_file(
    #     BytesIO(excel_bytes),
    #     as_attachment=True,
    #     download_name=filename,
    #     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    # )

    send_report_email(email_attachments)
    return Response("Report sent", status=200)
