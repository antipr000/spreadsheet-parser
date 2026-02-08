import openpyxl
import dotenv
from openpyxl.utils.cell import range_boundaries
from extractors.chart import ChartExtractor
from extractors.table import TableExtractor


dotenv.load_dotenv()


def ref_formula(obj):
    # returns something like "'Sheet1'!$B$2:$B$10"
    return getattr(getattr(obj, "numRef", None), "f", None) or getattr(
        getattr(obj, "strRef", None), "f", None
    )


def safe_title(chart):
    # title can be rich text; keep it best-effort
    try:
        if chart.title is None:
            return None
        # often chart.title.tx.rich.p[0].r[0].t or similar
        tx = getattr(chart.title, "tx", None)
        rich = getattr(tx, "rich", None)
        if rich and rich.p and rich.p[0].r and rich.p[0].r[0].t:
            return rich.p[0].r[0].t
    except Exception:
        pass
    return None


def cells_from_a1_range(wb, a1):
    # a1: "'Sheet1'!$B$2:$B$10" or "Sheet1!B2:B10"
    sheet_part, rng = a1.split("!")
    sheet_name = sheet_part.strip("'")
    ws2 = wb[sheet_name]
    min_col, min_row, max_col, max_row = range_boundaries(rng.replace("$", ""))
    out = []
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            out.append(ws2.cell(row=r, column=c).value)
    return out


def parse_excel(file_path):
    workbook = openpyxl.load_workbook(
        file_path,
        data_only=False,
        read_only=False,
        keep_links=True,
        keep_vba=True,
        rich_text=True,
    )
    print(workbook.sheetnames)

    sheet = workbook["Sheet_8"]

    table_extractor = TableExtractor()
    tables = table_extractor.extract(sheet)
    if tables:
        print(f"\n--- Tables in 'Charts' ---")
        for table in tables:
            print("--------------------------------")
            print("Heading:")
            for cell in table.heading:
                print(cell.value)
            print("Footer:")
            for cell in table.footer:
                print(cell.value)


if __name__ == "__main__":
    parse_excel("master2.xlsx")
