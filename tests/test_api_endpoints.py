import copy
from io import BytesIO

import api_endpoints
import pytest
from flask import Flask
from flask.testing import FlaskClient
from main import create_app
from openpyxl import load_workbook

VALID_PAYLOAD = {
    "regnskabsaar": "2026",
    "institution/afdeling": "Testinstitution",
    "kontaktperson": "Test Testesen",
    "telefon": "00000000",
    "dato": "01-01-2026",
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
        "nyanskaffet_anlaeg100000": [],
        "anskaffet_anlaeg": "Ja",
        "nyanskaffet_anlaeg": [
            {
                "nyanskaffet_a_beskrivelse": "Testpumpe",
                "nyanskaffet_a_bilagsnr": "1000000000",
                "nyanskaffet_a_pris": "32500",
                "nyanskaffet_a_aargang": "2026",
            },
        ],
    },
    "inventar": {
        "anskaffet_inventar_100000": "Nej",
        "nyanskaffet_inventar_100000": [],
        "nyanskaffet_inventar100000": [],
        "anskaffet_inventar": "Nej",
        "nyanskaffet_inventar": [],
    },
    "leasingaftaler": {
        "indgaaet_leasingaftaler": "Nej",
        "indgaaede_leasingaftaler": [],
        "aftalekontrakter": [],
        "opsagt_leasingaftaler": "Nej",
        "opsagte_leasingaftaler": [],
        "andre_bemaerkninger": "Ingen bemærkninger",
    },
}


@pytest.fixture()
def payload() -> dict:
    return copy.deepcopy(VALID_PAYLOAD)


@pytest.fixture()
def sent_emails(monkeypatch: pytest.MonkeyPatch) -> list:
    sent = []
    monkeypatch.setattr(
        api_endpoints,
        "send_report_email",
        lambda attachments: sent.append(attachments),
    )
    return sent


@pytest.fixture()
def client() -> FlaskClient:
    app: Flask = create_app()
    app.config.update({"TESTING": True})
    return app.test_client()


def test_valid_report_sends_email_with_excel(
    client: FlaskClient,
    payload: dict,
    sent_emails: list,
) -> None:
    response = client.get("/api/xflow-anlaeg", json=payload)

    assert response.status_code == 200
    assert len(sent_emails) == 1

    attachments = sent_emails[0]
    assert len(attachments) == 1

    filename, content = attachments[0]
    assert filename == "Anlaegsindberetning_Testinstitution_2026.xlsx"

    workbook = load_workbook(BytesIO(content))
    assert workbook["Overblik"]["B4"].value == "Testinstitution"
    assert workbook["Tekniske anlæg"]["A13"].value == "Testpumpe"


def test_contracts_are_attached(
    client: FlaskClient,
    payload: dict,
    sent_emails: list,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload["leasingaftaler"]["aftalekontrakter"] = [
        {"url": "https://example.invalid/kontrakt.pdf", "title": "kontrakt.pdf"},
    ]
    monkeypatch.setattr(
        api_endpoints,
        "download_contracts",
        lambda _payload: [
            {"filename": "kontrakt.pdf", "mime_type": "application/pdf", "content": b"%PDF-test"},
        ],
    )

    response = client.get("/api/xflow-anlaeg", json=payload)

    assert response.status_code == 200
    assert sent_emails[0][1] == ("kontrakt.pdf", b"%PDF-test")


def test_non_json_body_is_rejected(client: FlaskClient, sent_emails: list) -> None:
    response = client.get("/api/xflow-anlaeg", data="ikke json")

    assert response.status_code == 400
    assert sent_emails == []


def test_missing_required_field_is_rejected(
    client: FlaskClient,
    payload: dict,
    sent_emails: list,
) -> None:
    del payload["institution/afdeling"]

    response = client.get("/api/xflow-anlaeg", json=payload)

    assert response.status_code == 400
    assert b"institution/afdeling" in response.data
    assert sent_emails == []


def test_section_list_with_wrong_type_is_rejected(
    client: FlaskClient,
    payload: dict,
    sent_emails: list,
) -> None:
    payload["inventar"]["nyanskaffet_inventar"] = "ikke en liste"

    response = client.get("/api/xflow-anlaeg", json=payload)

    assert response.status_code == 400
    assert sent_emails == []