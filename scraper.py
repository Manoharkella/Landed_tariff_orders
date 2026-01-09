import pdfplumber
import json
import os
import shutil
import stat
from collections import defaultdict


def remove_readonly(func, path, _):
    os.chmod(path, stat.S_IWRITE)
    func(path)


def ensure_unique_headers(headers):
    counts = defaultdict(int)
    unique_headers = []

    for h in headers:
        h_str = str(h).strip().replace("\n", " ") if h else "Column"
        if counts[h_str] > 0:
            unique_headers.append(f"{h_str}_{counts[h_str]}")
        else:
            unique_headers.append(h_str)
        counts[h_str] += 1

    return unique_headers


def get_nearest_text_heading(page, table_top, max_distance=80):
    """
    Get nearest full text line immediately above the table
    (MOST reliable method for PDFs)
    """
    try:
        lines = page.extract_text_lines()
        candidates = []

        for line in lines:
            distance = table_top - line["bottom"]
            if 0 < distance <= max_distance:
                text = line["text"].strip()
                if text and len(text) > 3:
                    candidates.append((distance, text))

        if not candidates:
            return ""

        candidates.sort(key=lambda x: x[0])
        return candidates[0][1]

    except Exception:
        return ""


def scrape_pdf_tables_to_jsonl():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_root = os.path.join(base_dir, "Download")
    output_root = os.path.join(base_dir, "Extraction")

    if not os.path.exists(input_root):
        print(f"Error: Folder not found -> {input_root}")
        return

    if os.path.exists(output_root):
        shutil.rmtree(output_root, onerror=remove_readonly)

    os.makedirs(output_root, exist_ok=True)

    for root, _, files in os.walk(input_root):
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        if not pdf_files:
            continue

        rel_path = os.path.relpath(root, input_root)
        output_dir = os.path.join(output_root, rel_path)
        os.makedirs(output_dir, exist_ok=True)

        for pdf_file in sorted(pdf_files):
            pdf_path = os.path.join(root, pdf_file)
            output_path = os.path.join(
                output_dir, os.path.splitext(pdf_file)[0] + ".jsonl"
            )

            print(f"\nProcessing: {pdf_path}")

            with pdfplumber.open(pdf_path) as pdf, open(output_path, "w", encoding="utf-8") as f_out:

                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.find_tables()
                    if not tables:
                        continue

                    tables.sort(key=lambda t: t.bbox[1])
                    prev_bottom = 0

                    for table_index, table in enumerate(tables, start=1):
                        data = table.extract()
                        if not data or len(data) < 2:
                            continue

                        _, table_top, _, table_bottom = table.bbox
                        search_top = max(prev_bottom, table_top - 80)

                        table_heading = get_nearest_text_heading(
                            page,
                            table_top
                        )

                        prev_bottom = table_bottom

                        headers = ensure_unique_headers(data[0])

                        rows = []
                        for row in data[1:]:
                            row_obj = {
                                headers[i]: (
                                    row[i].strip()
                                    if i < len(row) and isinstance(row[i], str)
                                    else row[i]
                                )
                                for i in range(len(headers))
                            }
                            rows.append(row_obj)

                        record = {
                            "document_name": pdf_file,
                            "page_number": page_num,
                            "table_index": table_index,
                            "table_heading": table_heading,
                            "headers": headers,
                            "rows": rows
                        }

                        f_out.write(json.dumps(record, ensure_ascii=False) + "\n")

            print("✔ Completed")


if __name__ == "__main__":
    scrape_pdf_tables_to_jsonl()
