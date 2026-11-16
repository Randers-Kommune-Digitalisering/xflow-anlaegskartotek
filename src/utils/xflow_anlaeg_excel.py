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

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Nyanskaffede ejendomme",
        data.get("anskaffet_fast_ejendom"),
        data.get("nyanskaffet_ejendom", []),
        [
            ("Beskrivelse", "nyanskaffet_e_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_e_bilagsnr"),
            ("Dato", "nyanskaffet_e_dato"),
            ("Pris", "nyanskaffet_e_pris"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Afhændede ejendomme",
        data.get("afhaendet_fast_ejendom"),
        data.get("nyafhaendet_ejendom", []),
        [
            ("Beskrivelse", "nyafhaendet_e_beskrivelse"),
            ("Bilagsnr.", "nyafhaendet_e_bilagsnr"),
            ("Dato", "nyafhaendet_e_dato"),
            ("Pris", "nyafhaendet_e_pris"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Foretagne forbedringer",
        data.get("foretaget_forbedringer"),
        data.get("forbedringer", []),
        [
            ("Beskrivelse", "forbedring_beskrivelse"),
            ("Bilagsnr.", "forbedring_bilagsnr"),
            ("Dato", "forbedring_dato"),
            ("Pris", "forbedring_pris"),
        ],
    )

    # Tekniske anlæg
    ws = workbook.create_sheet("Tekniske anlæg")
    data = payload["tekniske_anlaeg"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Nyanskaffede anlæg over 100.000 kr.",
        data.get("anskaffet_anlaeg_100000"),
        data.get("nyanskaffet_anlaeg_100000", []),
        [
            ("Beskrivelse", "nyanskaffet_a1_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_a1_bilagsnr"),
            ("Pris", "nyanskaffet_a1_pris"),
            ("Årgang", "nyanskaffet_a1_aarang"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Nyanskaffede anlæg under 100.000 kr.",
        data.get("anskaffet_anlaeg"),
        data.get("nyanskaffet_anlaeg", []),
        [
            ("Beskrivelse", "nyanskaffet_a_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_a_bilagsnr"),
            ("Pris", "nyanskaffet_a_pris"),
            ("Årgang", "nyanskaffet_a_aargang"),
        ],
    )

    # Inventar
    ws = workbook.create_sheet("Inventar")
    data = payload["inventar"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Nyanskaffet inventar over 100.000 kr.",
        data.get("anskaffet_inventar_100000"),
        data.get("nyanskaffet_inventar_100000", []),
        [
            ("Beskrivelse", "nyanskaffet_i1_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_i1_bilagsnr"),
            ("Pris", "nyanskaffet_i1_pris"),
            ("Årgang", "nyanskaffet_i1_aargang"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Nyanskaffet inventar under 100.000 kr.",
        data.get("anskaffet_inventar"),
        data.get("nyanskaffet_inventar", []),
        [
            ("Beskrivelse", "nyanskaffet_i_beskrivelse"),
            ("Bilagsnr.", "nyanskaffet_i_bilagsnr"),
            ("Pris", "nyanskaffet_i_pris"),
            ("Årgang", "nyanskaffet_i_aargang"),
        ],
    )

    # Leasingaftaler
    ws = workbook.create_sheet("Leasingaftaler")
    data = payload["leasingaftaler"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Indgåede leasingaftaler",
        data.get("indgaaet_leasingaftaler"),
        data.get("indgaaede_leasingaftaler", []),
        [
            ("Beskrivelse", "indgaaet_beskrivelse"),
            ("Bemærkninger", "indgaaet_bemaerkninger"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Opsagte leasingaftaler",
        data.get("opsagt_leasingaftaler"),
        data.get("opsagte_leasingaftaler", []),
        [
            ("Beskrivelse", "opsagt_beskrivelse"),
            ("Bemærkninger", "opsagt_bemaerkninger"),
        ],
    )

    # Attached contracts
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


def write_section(
        ws,
        start_row: int,
        title: str,
        answer: str,
        records: list,
        columns: list,
) -> int:
    """
    Write a section in the worksheet.
    
    :worksheet ws: The worksheet to write to.
    :start_row: The starting row for the section.
    :title: The title of the section.
    :answer: The answer text for the section.
    :records: A list of records to include in the section.
    :columns: A list of column definitions for the records.
    """

    ws.cell(start_row, 1, title)
    ws.cell(start_row, 1).font = TITLE_FONT

    ws.cell(start_row + 1, 1, "Svar")
    ws.cell(start_row + 1, 2, answer or "")

    write_table(
        ws,
        start_row + 3,
        records,
        columns,
    )
    return start_row + max(8, len(records) + 4)


def write_table(ws, start_row: int, records: list, columns: list) -> None:
    """
    Write a table with headers and records.

    :worksheet ws: The worksheet to write to.
    :start_row: The starting row for the table.
    :records: A list of records to include in the table.
    :columns: A list of column definitions for the records.
    """

    headers = [label for label, _ in columns]

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(start_row, col, header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left")

    if not records:
        ws.cell(start_row + 1, 1, "Ingen registreringer")
        return

    for row_index, record in enumerate(records, start=start_row + 1):
        for col_index, (_, key) in enumerate(columns, start=1):
            ws.cell(
                row_index,
                col_index,
                record.get(key, ""),
            )


def auto_adjust_columns(ws) -> None:
    """
    Automatically adjust column widths
    :worksheet ws: The worksheet to adjust columns for.
    """

    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value)),
                )

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 3, 15),
            60,
        )
