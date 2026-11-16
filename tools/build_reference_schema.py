from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import load_workbook


def main() -> None:
    source = Path(sys.argv[1])
    destination = Path(sys.argv[2])
    sheet = load_workbook(source, read_only=True, data_only=False).active
    headers = [sheet.cell(1, column).value for column in range(4, sheet.max_column + 1)]
    rows = [list(row[:3]) for row in sheet.iter_rows(min_row=2, values_only=True)]
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps({"version": 1, "sheet": sheet.title, "headers": headers, "rows": rows}, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
