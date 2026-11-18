"""Generate Excel reports for XFlow asset reports."""

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

HEADER_FILL = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid",
)

HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(size=14, bold=True, color="1F4E78")
LABEL_FONT = Font(color="1F4E78", bold=True)


def create_xflow_anlaeg_excel(payload: dict) -> bytes:
    """
    Generate an Excel workbook from the XFlow report.

    :param payload: The data to be included in the Excel report.
    :return: The generated Excel file as bytes.
    """
    workbook = Workbook()

    # Overview
    ws = workbook.active
    ws.title = "Overblik"

    ws["A1"] = "Registrering af fysiske aktiver til anlægskartotek"
    ws["A1"].font = TITLE_FONT

    overview = [
        ("Regnskabsår", payload.get("regnskabsaar")),
        ("Institution/afdeling", payload.get("institution/afdeling")),
        ("Kontaktperson", payload.get("kontaktperson")),
        ("Telefon", payload.get("telefon")),
        ("Dato", payload.get("dato")),
    ]

    for row, (label, value) in enumerate(overview, start=3):
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = LABEL_FONT
        ws.cell(row=row, column=2, value=value)

    # Grunde og bygninger
    ws = workbook.create_sheet("Grunde og bygninger")
    data = payload["grunde_og_bygninger"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Anskaffede ejendomme",
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
        "Væsentlige forbedringer/renoveringer",
        data.get("foretaget_forbedringer"),
        data.get("forbedringer", []),
        [
            ("Beskrivelse", "forbedring_beskrivelse"),
            ("Bilagsnr.", "forbedring_bilagsnr"),
            ("Dato", "forbedring_dato"),
            ("Pris", "forbedring_pris"),
        ],
    )

    if start_row == 1:
        workbook.remove(ws)

    # Tekniske anlæg
    ws = workbook.create_sheet("Tekniske anlæg, maskiner, større specialudstyr og transportmidler")
    data = payload["tekniske_anlaeg"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Anskaffede anlæg, maskiner, større specialudstyr og transportmidler",
        data.get("anskaffet_anlaeg"),
        data.get("nyanskaffet_anlaeg", []),
        [
            ("Beskrivelse", "anskaffet_a_beskrivelse"),
            ("Bilagsnr.", "anskaffet_a_bilagsnr"),
            ("Pris", "anskaffet_a_pris"),
            ("Årgang", "anskaffet_a_aarang"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Afhaendede anlæg, maskiner, større specialudstyr og transportmidler",
        data.get("afhaendet_anlaeg"),
        data.get("nyafhaendet_anlaeg", []),
        [
            ("Beskrivelse", "afhaendet_a_beskrivelse"),
            ("Bilagsnr.", "afhaendet_a_bilagsnr"),
            ("Pris", "afhaendet_a_pris"),
            ("Årgang", "afhaendet_a_aargang"),
        ],
    )

    if start_row == 1:
        workbook.remove(ws)

    # Inventar
    ws = workbook.create_sheet("Inventar, herunder computere og andet IT-udstyr")
    data = payload["inventar"]

    start_row = 1
    start_row = write_section(
        ws,
        start_row,
        "Anskaffet inventar, computere og andet IT-udstyr",
        data.get("anskaffet_inventar"),
        data.get("nyanskaffet_inventar", []),
        [
            ("Beskrivelse", "anskaffet_i_beskrivelse"),
            ("Bilagsnr.", "anskaffet_i_bilagsnr"),
            ("Pris", "anskaffet_i_pris"),
            ("Årgang", "anskaffet_i_aargang"),
        ],
    )

    start_row = write_section(
        ws,
        start_row,
        "Afhændet inventar, computere og andet IT-udstyr",
        data.get("afhaendet_inventar"),
        data.get("nyafhaendet_inventar", []),
        [
            ("Beskrivelse", "afhaendet_i_beskrivelse"),
            ("Bilagsnr.", "afhaendet_i_bilagsnr"),
            ("Pris", "afhaendet_i_pris"),
            ("Årgang", "afhaendet_i_aargang"),
        ],
    )
    if start_row == 1:
        workbook.remove(ws)

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

    ws.cell(
        start_row,
        1,
        "Andre bemærkninger",
    )

    ws.cell(
        start_row,
        1,
    ).font = TITLE_FONT

    ws.cell(
        start_row + 1,
        1,
        data.get("andre_bemaerkninger", ""),
    )

    if (
        data.get("indgaaet_leasingaftaler") != "Ja"
        and data.get("opsagt_leasingaftaler") != "Ja"
        and not str(data.get("andre_bemaerkninger") or "").strip()
    ):
        workbook.remove(ws)

    # Format all worksheets
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A3"
        sheet.sheet_view.showGridLines = True
        auto_adjust_columns(sheet)

    # Return Excel as bytes
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()


def write_section(
        ws: Worksheet,
        start_row: int,
        title: str,
        answer: str,
        records: list,
        columns: list,
) -> int:
    """
    Write a section in the worksheet.

    :param ws: The worksheet to write to.
    :param start_row: The starting row for the section.
    :param title: The title of the section.
    :param answer: The answer text for the section.
    :param records: A list of records to include in the section.
    :param columns: A list of column definitions for the records.
    :return: The row number after the section.
    """
    if answer != "Ja":
        return start_row

    title_cell = ws.cell(start_row, 1)

    if " over " in title or " under " in title:
        word = "over" if " over " in title else "under"
        before, after = title.split(word, 1)
        normal = InlineFont(sz=14, b=True, color="1F4E78")
        underlined = InlineFont(sz=14, b=True, u="single", color="1F4E78")
        title_cell.value = CellRichText(
            TextBlock(normal, before),
            TextBlock(underlined, word),
            TextBlock(normal, after),
        )
    else:
        title_cell.value = title
        title_cell.font = TITLE_FONT

    write_table(
        ws,
        start_row + 1,
        records,
        columns,
    )
    return start_row + max(7, len(records) + 3)


def write_table(ws: Worksheet, start_row: int, records: list, columns: list) -> None:
    """
    Write a table with headers and records.

    :param ws: The worksheet to write to.
    :param start_row: The starting row for the table.
    :param records: A list of records to include in the table.
    :param columns: A list of column definitions for the records.
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


def auto_adjust_columns(ws: Worksheet) -> None:
    """
    Automatically adjust column widths.

    :param ws: The worksheet to adjust columns for.
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
