"""
Receive XFlox asset and leasing reports
"""
import logging
import os

from flask import Blueprint, Response, current_app, request, send_file

from io import BytesIO
from pathlib import PurePath
from urllib.parse import urlsplit
from zipfile import ZIP_DEFLATED, ZipFile

import requests

from src.utils.xflow_anlaeg_excel import create_xflow_anlaeg_excel

api = Blueprint("xflow_anlaeg", __name__, url_prefix="/api")
logger = logging.getLogger(__name__)


@api.route("/xflow-anlaeg", methods=["GET"])
def get_xflow_anlaeg_data() -> Response:
    """
    Receive an asset report submitted through XFLow
    """

    # Allow mock payload during local development
    if current_app.debug and os.getenv("XFLOW_ANLAEG_MOCK") == "1":
        payload = {
            "regnskabsaar": "2026",
            "instutition/afdeling": "Testinstitution",
            "kontaktperson": "Testperson",
            "telefon": "12345678",
            "dato": "11-12-2026",
            "grunde_og_bygninger": {
                "anskaffet_fast_ejendom": "Nej",
                "nyanskaffet_ejendom": [],
                "afhaendet_fast_ejendom": "Nej",
                "nyafhaendet_ejendom": [],
                "foretaget_forbedringer": "Nej",
                "forbedringer": []
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
                        "nyanskaffet_a_aargang": "2026"
                    }
                ]
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
                        "nyanskaffet_i_aargang": "2026"
                    }
                ]
            },
            "leasingaftaler": {
                "indgaaet_leasingaftaler": "Nej",
                "indgaaede_leasingaftaler": [],
                "aftalekontrakter": [],
                "opsagt_leasingaftaler": "Nej",
                "opsagte_leasingaftaler": [],
                "andre_bemaerkninger": "Der er ikke indgået eller opsagt leasingaftaler i regnskabsåret. Der er derfor ingen leasingkontrakt at vedhæfte."
            },
            "process_id": "mock_process_id"
        }

        logger.info("Using XFlow asset mock payload")

    else:
        payload = request.get_json(silent=True)

        if not isinstance(payload, dict):
            return Response(
                "Invalid payload: expected a JSON object", status=400
            )

    # Validate required top-level fields
    required_fields = [
        "regnskabsaar", 
        "instutition/afdeling",
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
                f"Invalid payload: missing required field '{field}'", status=400
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
                f"Invalid payload: '{section}' must be a JSON object", status=400
            )

    # Validate repeating fields
    array_fields = {
        "grunde_og_bygninger": [
            "nyanskaffet_ejendom",
            "nyafhaendet_ejendom",
            "forbedringer",
        ],
        "tekniske_anlaeg": [
            "nyanskaffet_anlaeg100000",
            "nyanskaffet_anlaeg",
        ],
        "inventar": [
            "nyanskaffet_inventar100000",
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
                    f"Invalid payload: '{section}.{field}' must me an array.", status=400
                )

    # Extract information
    regnskabsaar = payload["regnskabsaar"]
    insitution = payload["institution/afdeling"]

    leasingaftaler = payload["leasingaftaler"]
    contracts = leasingaftaler["aftalekontrakter"]

    logger.info(
        "XFlow asset report recived: year=%s, intitution=%s",
        regnskabsaar,
        insitution
    )

    logger.info(
        "Number of attached leasing contracts: %s",
        len(contracts)
    )

    try:
        excel_bytes = create_xflow_anlaeg_excel(payload)
    except (KeyError, TypeError, ValueError):
        logger.exception("Failed to generate XFlow Excel report")
        return Response(
            "Excel report could not be generated", status=500
        )

    filename = (
        f"Anlaegsindberetning_"
        f"{payload['institution/afdeling']}_"
        f"{payload['regnskabsaar']}.xlsx"
    )

    logger.info("XFlow Excel report generated: filename=%s", filename)

    token = os.getenv("XFLOW_PUBLIC_API_TOKEN")
    headers = {"publicApiToken": token} if token else {}
    allowed_host = os.environ["XFLOW_DOCUMENT_HOST"]

    zip_buffer = BytesIO()

    with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("Anlaegsinberetning.xlsx", excel_bytes)

        for index, contract in enumerate(contracts, start=1):
            url = contract["url"]
            parsed_url = urlsplit(url)

            if(
                parsed_url.scheme != "https"
                or parsed_url.hostname != allowed_host
                or parsed_url.username is not None
                or parsed_url.password is not None
            ):
                return Response("Invalid document URL", status=400)

            title = contract["title"]
            filename = PurePath(title.replace("\\", "/")).name
            if not filename or filename in {".", ".."}:
                return Response("Invalid document filename", status=400)

            try:
                with requests.get(
                    url,
                    headers=headers,
                    timeout=(5, 30),
                    allow_redirects=False,
                    stream=True,
                ) as response:
                    response.raise_for_status()
                    if 300 <= response.status_code < 400:
                        return Response("Document redirect not allowed", status=502)

                    content = response.content
            except requests.RequestException:
                logger.exception("Could not retrieve XFlow document")
                return Response("Could not retrieve attached document", status=502)

            archive.writestr(f"kontrakter/{index}_{filename}", content)

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="Anlaegsindberetning.zip"
    )