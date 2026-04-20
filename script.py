import re
from pathlib import Path

import pdfplumber
from PyPDF2 import PdfReader, PdfWriter

INPUT_PDF = "input.pdf"
OUTPUT_FOLDER = "output_pdfs"
PURCHASING_FOLDER = None


def _find_standard_candidates(text):
    return re.findall(r"\b\d?[A-Z]{2,3}\d{4,6}[A-Z]?\b", text.upper())


def _deinterleave_after_label(compact_tail, label):
    """
    OCR can interleave the part number into label text.
    Example:
      PART NUMBER: 3PS03560M
    becomes:
      PA3RTPSN0UM35BE6R0M
    """
    j = 0
    payload_chars = []
    for ch in compact_tail:
        if j < len(label) and ch == label[j]:
            j += 1
        else:
            payload_chars.append(ch)

    if j < len(label) - 1:
        return None

    payload = "".join(payload_chars)
    candidates = _find_standard_candidates(payload)
    if candidates:
        return max(candidates, key=len)
    return None


def _has_interleaved_label(compact_tail, label):
    j = 0
    for ch in compact_tail:
        if j < len(label) and ch == label[j]:
            j += 1
    return j >= len(label) - 1


def extract_part_number(text):
    text = (text or "").replace("\n", " ")
    text_upper = text.upper()

    rev_region_match = re.search(
        r"REV\s*:.*?DESCRIPTION\s*:",
        text_upper,
        flags=re.DOTALL,
    )
    region = rev_region_match.group(0) if rev_region_match else text_upper
    compact = re.sub(r"[^A-Z0-9]+", "", region)

    pre_desc = compact.split("DESCRIPTION", 1)[0]
    of_match = None
    for m in re.finditer(r"OF\d+", pre_desc):
        of_match = m
    compact_tail = pre_desc[of_match.end():] if of_match else pre_desc

    for label in ("PARTNUMBER", "PARTSNUMBER"):
        deinterleaved = _deinterleave_after_label(compact_tail, label)
        if deinterleaved:
            return deinterleaved

    # If the label is present but OCR noise prevents standard token extraction,
    # prefer the page's structured drawing identifier over BOM row part numbers.
    label_found = any(
        _has_interleaved_label(compact_tail, label)
        for label in ("PARTNUMBER", "PARTSNUMBER")
    )
    if label_found:
        structured = re.findall(
            r"\b[A-Z]{1,4}-\d{2,6}-[A-Z0-9-]+-\d+\.?\d*[A-Z]*\b",
            text_upper,
        )
        if structured:
            return max(structured, key=len)

    region_candidates = _find_standard_candidates(compact)
    if region_candidates:
        return region_candidates[0]

    direct_candidates = _find_standard_candidates(text_upper)
    if direct_candidates:
        ps_priority = [c for c in direct_candidates if re.match(r"^\d?PS\d{4,6}[A-Z]?$", c)]
        if ps_priority:
            return ps_priority[0]
        return direct_candidates[0]

    # Keep generic structured IDs as a last resort because values like
    # "PS-245-UC6K-178" can be page-set identifiers repeated on every page.
    structured = re.findall(
        r"\b[A-Z]{1,4}-\d{2,6}-[A-Z0-9-]+-\d+\.?\d*[A-Z]*\b",
        text_upper,
    )
    if structured:
        return max(structured, key=len)

    return None


def process_pdf(input_pdf, output_folder, purchasing_folder=None):
    input_pdf = Path(input_pdf)
    output_dir = Path(output_folder)
    output_dir.mkdir(parents=True, exist_ok=True)
    purchasing_dir = Path(purchasing_folder) if purchasing_folder else None
    if purchasing_dir:
        purchasing_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))
    seen_part_numbers = set()
    imported_count = 0
    duplicate_in_run_count = 0
    already_exists_count = 0
    no_part_number_count = 0
    purchasing_folder_missing_count = 0
    details = []

    with pdfplumber.open(str(input_pdf)) as pdf:
        for i, page in enumerate(pdf.pages):
            page_num = i + 1
            text = page.extract_text() or ""
            part_number = extract_part_number(text)

            if not part_number:
                message = f"Skip page {page_num}: no part number found"
                print(message)
                no_part_number_count += 1
                details.append(
                    {
                        "page": page_num,
                        "part_number": None,
                        "status": "not_imported",
                        "reason": "no_part_number",
                        "output_path": None,
                        "message": message,
                    }
                )
                continue

            if part_number in seen_part_numbers:
                message = f"Skip page {page_num}: duplicate part number {part_number}"
                print(message)
                duplicate_in_run_count += 1
                details.append(
                    {
                        "page": page_num,
                        "part_number": part_number,
                        "status": "not_imported",
                        "reason": "duplicate_in_input",
                        "output_path": None,
                        "message": message,
                    }
                )
                continue
            seen_part_numbers.add(part_number)

            safe_part_number = re.sub(r'[\\/*?:"<>|]', "_", part_number)
            is_purchasing_part = safe_part_number.startswith("0")
            target_dir = output_dir
            if is_purchasing_part:
                if purchasing_dir is None:
                    message = (
                        f"Skip page {page_num}: purchasing folder not selected for "
                        f"part number {part_number}"
                    )
                    print(message)
                    purchasing_folder_missing_count += 1
                    details.append(
                        {
                            "page": page_num,
                            "part_number": part_number,
                            "status": "not_imported",
                            "reason": "purchasing_folder_missing",
                            "output_path": None,
                            "message": message,
                        }
                    )
                    continue
                target_dir = purchasing_dir

            writer = PdfWriter()
            writer.add_page(reader.pages[i])

            output_path = target_dir / f"{safe_part_number}.pdf"
            if output_path.exists():
                message = f"Skip page {page_num}: already exists {output_path}"
                print(message)
                already_exists_count += 1
                details.append(
                    {
                        "page": page_num,
                        "part_number": part_number,
                        "status": "not_imported",
                        "reason": "already_exists",
                        "output_path": str(output_path),
                        "message": message,
                    }
                )
                continue

            with open(output_path, "wb") as f:
                writer.write(f)

            message = f"Saved: {output_path} (from page {page_num})"
            print(message)
            imported_count += 1
            details.append(
                {
                    "page": page_num,
                    "part_number": part_number,
                    "status": "imported",
                    "reason": "saved",
                    "output_path": str(output_path),
                    "message": message,
                }
            )

    not_imported_count = (
        duplicate_in_run_count
        + already_exists_count
        + no_part_number_count
        + purchasing_folder_missing_count
    )

    print(f"Generated {imported_count} unique PDF(s).")
    print(f"Skipped duplicates in input: {duplicate_in_run_count}")
    print(f"Skipped existing files: {already_exists_count}")
    print(f"Skipped pages with no part number: {no_part_number_count}")
    print(f"Skipped purchasing parts (folder not selected): {purchasing_folder_missing_count}")
    print(f"Not imported total: {not_imported_count}")

    return {
        "input_pdf": str(input_pdf),
        "output_folder": str(output_dir),
        "purchasing_folder": str(purchasing_dir) if purchasing_dir else "",
        "total_pages": len(reader.pages),
        "imported_count": imported_count,
        "not_imported_count": not_imported_count,
        "duplicate_in_input_count": duplicate_in_run_count,
        "already_exists_count": already_exists_count,
        "no_part_number_count": no_part_number_count,
        "purchasing_folder_missing_count": purchasing_folder_missing_count,
        "details": details,
    }


def main():
    process_pdf(INPUT_PDF, OUTPUT_FOLDER, PURCHASING_FOLDER)


if __name__ == "__main__":
    main()
