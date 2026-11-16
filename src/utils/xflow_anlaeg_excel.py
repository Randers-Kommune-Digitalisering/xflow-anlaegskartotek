"""Generate Excel reports for XFlow asset reports."""

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(size=14, bold=True, color="1F4E78")

def create_xflow_anlaeg_excel(payload:dict) -> bytes:
    """
    Generate an Excel workbook from the XFlow report.

    :param payload: The data to be included in the Excel report.
    :return: The generated Excel file as bytes.
    """

    workbook = Workbook()

    # Overview
    ws = workbook.active
    ws.title = "Overblik"

    ws["A1"] = "Anlægsindberetning"
    ws["A1"].font = TITLE_FONT

    overview = [
        ("Regnskabsår", payload.get("regnskabsaar")),
        ("Institution/afdeling", payload.get("institution/afdeling")),
        ("Kontaktperson", payload.get("kontaktperson")),
        ("Telefon", payload.get("telefon")),
        ("Dato", payload.get("dato")),
    ]

    for row, (label, value) in enumerate(overview, start=3):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=value)

    # Grunde og bygninger
    ws = workbook.create_sheet("Grunde og bygninger")
    data = payload["grunde_og_bygninger"]

    write_section(
        ws,
        1,
        "Nyanskaffede ejendomme",
        data.get("anskaffet_fast_ejendom"),
        data.get("nyanskaffet_ejendom", []),
        [
            ("Beskrivelse", "nyanskaffet_e_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_e_bilagsnr"),
            ("Type", "nyanskaffet_e_type"),
            ("Dato", "nyanskaffet_e_dato"),
            ("Pris", "nyanskaffet_e_pris"),
        ],
    )

    write_section(
        ws,
        9,
        "Afhændede ejendomme",
        data.get("afhaendet_fast_ejendom"),
        data.get("nyafhaendet_ejendom", []),
        [
            ("Beskrivelse", "nyafhaendet_e_beskrivelse"),
            ("Bilagsnr.", "nyafhaendet_e_bilagsnr"),
            ("Type", "nyafhaendet_e_type"),
            ("Dato", "nyafhaendet_e_dato"),
            ("Pris", "nyafhaendet_e_pris"),
        ],
    )

    write_section(
        ws,
        17,
        "Foretagne forbedringer",
        data.get("foretaget_forbedringer"),
        data.get("forbedringer_e", []),
        [
            ("Beskrivelse", "forbedring_beskrivelse"),
            ("Bilagsnr.", "forbedring_bilagsnr"),
            ("Type", "forbedring_type"),
            ("Dato", "forbedring_dato"),
            ("Pris", "forbedring_pris"),
        ],
    )

    # Tekniske anlæg
    ws = workbook.create_sheet("Tekniske anlæg")
    data = payload["tekniske_anlaeg"]

    write_section(
        ws,
        1,
        "Nyanskaffede anlæg over 100.000 kr.",
        data.get("anskaffet_anlaeg_100000"),
        data.get("nyanskaffet_anlaeg_100000", []),
        [
            ("Beskrivelse", "nyanskaffet_a1_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_a1_bilagsnr"),
            ("Type", "nyanskaffet_a1_type"),
            ("Pris", "nyanskaffet_a1_pris"),
            ("Årgang", "nyanskaffet_a1_aarang"),
        ],
    )

    write_section(
        ws,
        9,
        "Nyanskaffede anlæg under 100.000 kr.",
        data.get("anskaffet_anlaeg"),
        data.get("nyanskaffet_anlaeg", []),
        [
            ("Beskrivelse", "nyanskaffet_a_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_a_bilagsnr"),
            ("Type", "nyanskaffet_a_type"),
            ("Pris", "nyanskaffet_a_pris"),
            ("Årgang", "nyanskaffet_a_aargang"),
        ],
    )

    # Inventar
    ws = workbook.create_sheet("Inventar")
    data = payload["inventar"]

    write_section(
        ws,
        1,
        "Nyanskaffet inventar over 100.000 kr.",
        data.get("anskaffet_inventar_100000"),
        data.get("nyanskaffet_inventar_100000", []),
        [
            ("Beskrivelse", "nyanskaffet_i1_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_i1_bilagsnr"),
            ("Type", "nyanskaffet_i1_type"),
            ("Pris", "nyanskaffet_i1_pris"),
            ("Årgang", "nyanskaffet_i1_aargang"),
        ],
    )

    write_section(
        ws,
        9,
        "Nyanskaffet inventar under 100.000 kr.",
        data.get("anskaffet_inventar"),
        data.get("nyanskaffet_inventar", []),
        [
            ("Beskrivelse", "nyanskaffet_i_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_i_bilagsnr"),
            ("Type", "nyanskaffet_i_type"),
            ("Pris", "nyanskaffet_i_pris"),
            ("Årgang", "nyanskaffet_i_aargang"),
        ],
    )

    # Leasingaftaler
    ws = workbook.create_sheet("Leasingaftaler")
    data = payload["leasingaftaler"]

    write_section(
        ws,
        1,
        "Indgåede leasingaftaler",
        data.get("indgaaet_leasingaftaler"),
        data.get("indgaaede_leasingaftaler", []),
        [
            ("Beskrivelse", "indgaaet_beskrivelse"),
            ("Type", "indgaaet_type"),
            ("Bemærkninger", "indgaaet_bemaerkninger"),
        ],
    )

    write_section(
        ws,
        9,
        "Opsagte leasingaftaler",
        data.get("opsagt_leasingaftaler"),
        data.get("opsagte_leasingaftaler", []),
        [
            ("Beskrivelse", "opsagt_beskrivelse"),
            ("Type", "opsagt_type"),
            ("Bemærkninger", "opsagt_bemaerkninger"),
        ],
    )

    # Attached contracts
    start_row = 17
    ws.cell(start_row, 1, "Vedhæftede leasingkontrakter")
    ws.cell(start_row, 1).font = TITLE_FONT

    contracts = data.get("aftalekontrakter", [])

    contract_columns = [
        ("Filnavn", "title"),
        ("Filtype", "mimeType"),
        ("Filstørrelse (bytes)", "size"),
        ("Checksum", "checksum"),
        ("Fil URL", "url"),
    ]

    write_table(
        ws,
        start_row + 2,
        contracts,
        contract_columns,
    )

    ws.cell(
        start_row + len(contracts) + 4,
        1,
        "Andre bemærkninger",
    )

    ws.cell(
        start_row + len(contracts) + 5,
        1,
        data.get("andre_bemaerkninger", ""),
    )

    # Format all worksheets
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A3"
        sheet.sheet_view.showGridLines = False
        auto_adjust_columns(sheet)

    # Return Excel as bytes
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()
