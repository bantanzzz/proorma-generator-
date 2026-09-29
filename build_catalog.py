import json
from pathlib import Path
import openpyxl

source = Path(__file__).parent / "prices lists"
result = {}
for workbook_path in sorted(source.glob("*.xlsx")):
    workbook = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    vendor = workbook_path.stem.replace("_", " ").replace("-", " ").strip()
    items = []
    for sheet in workbook.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        for header_index, row in enumerate(rows):
            headers = [str(value or "").strip().lower() for value in row]
            description_col = next((i for i, value in enumerate(headers) if value in {"description", "item", "product", "item name", "name"}), None)
            price_col = next((i for i, value in enumerate(headers) if "unit price" in value or value == "price"), None)
            if description_col is None or price_col is None:
                continue
            unit_col = next((i for i, value in enumerate(headers) if value == "unit"), None)
            for values in rows[header_index + 1:]:
                description = values[description_col] if description_col < len(values) else None
                price = values[price_col] if price_col < len(values) else None
                unit = values[unit_col] if unit_col is not None and unit_col < len(values) else ""
                try:
                    price = float(price)
                except (TypeError, ValueError):
                    continue
                if not description or str(description).strip().lower() in {"grand total", "total"} or price <= 0:
                    continue
                items.append({"description": str(description).strip(), "unit": str(unit or "").strip(), "price": price, "category": sheet.title})
            break
    if items:
        result[vendor] = items
output = Path(__file__).parent / "vendor_catalog.js"
output.write_text("window.PRICE_LISTS = " + json.dumps(result, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
print(f"Updated {output.name}: {sum(map(len, result.values()))} items across {len(result)} vendors")
