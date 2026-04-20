import re
from pathlib import Path

import pdfplumber
from PyPDF2 import PdfReader, PdfWriter

INPUT_PDF = "input.pdf"
OUTPUT_FOLDER = "output_pdfs"
PURCHASING_FOLDER = None
CONFIDENCE_THRESHOLD = 0.72
MANUAL_REVIEW_FOLDER_NAME = "_manual_review"


def _find_standard_candidates(text):
    # Support common part-number shapes:
    # - 1HS01357M
    # - 0HWG01020
    # - CHS8600RSF
    return re.findall(r"\b\d?[A-Z]{2,4}\d{4,6}[A-Z]{0,4}\b", text.upper())


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


def _capture_direct_label_candidates(text_upper):
    """
    Prefer explicit PART NUMBER label values when OCR keeps words in reading order.
    """
    label_patterns = (
        r"PART\s*NUMBER\s*:?\s*([A-Z0-9\-]{4,20})",
        r"PARTS\s*NUMBER\s*:?\s*([A-Z0-9\-]{4,20})",
    )
    results = []
    for pattern in label_patterns:
        for match in re.findall(pattern, text_upper):
            cleaned = re.sub(r"[^A-Z0-9\-]", "", match)
            results.extend(_find_standard_candidates(cleaned))
    return results


def _first_candidate_after_part_number_label(text_upper):
    for match in re.finditer(r"PART\s*NUMBER", text_upper):
        segment = text_upper[match.end() : match.end() + 220]
        candidates = _find_standard_candidates(segment)
        if candidates:
            return candidates[0]
    return None


def _capture_candidates_with_sources(text_upper):
    candidates = {}
    is_parts_list_page = "PARTS LIST" in text_upper

    def add(candidate, source):
        if not candidate:
            return
        candidates.setdefault(candidate, set()).add(source)

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
            # Keep this signal even on PARTS LIST pages. Those sheets often
            # contain the true title-block part number in OCR-interleaved form.
            add(deinterleaved, "deinterleaved_label")

    if not is_parts_list_page:
        for c in _capture_direct_label_candidates(text_upper):
            add(c, "direct_label")

    if not is_parts_list_page:
        first_after_label = _first_candidate_after_part_number_label(text_upper)
        if first_after_label:
            add(first_after_label, "first_after_label")

    for c in _find_standard_candidates(compact):
        add(c, "title_region")

    direct_candidates = _find_standard_candidates(text_upper)
    for c in direct_candidates:
        add(c, "page_text")
        if re.match(r"^\d?PS\d{4,6}[A-Z]?$", c):
            add(c, "ps_priority")

    structured = re.findall(
        r"\b[A-Z]{1,4}-\d{2,6}-[A-Z0-9-]+-\d+\.?\d*[A-Z]*\b",
        text_upper,
    )
    for c in structured:
        add(c, "structured")
        if is_parts_list_page:
            # On PARTS LIST sheets, hyphenated drawing IDs in title blocks
            # are often the intended part number for that page.
            add(c, "parts_list_structured")

    return candidates


def _score_candidate(candidate, sources, text_upper, prefer_parts_list_structured=False):
    score = 0
    source_weights = {
        "deinterleaved_label": 10,
        "direct_label": 10,
        "first_after_label": 7,
        "title_region": 4,
        "page_text": 3,
        "ps_priority": 2,
        "structured": 1,
        "parts_list_structured": 8 if prefer_parts_list_structured else 0,
    }
    for source in sources:
        score += source_weights.get(source, 0)

    if candidate in text_upper:
        score += 2

    part_label_context = re.search(
        r"(PART\s*NUMBER|PARTS\s*NUMBER).{0,120}" + re.escape(candidate),
        text_upper,
        flags=re.DOTALL,
    )
    if part_label_context:
        score += 3

    if re.search(re.escape(candidate) + r"\s*-\s*", text_upper):
        score += 1

    if re.search(r"\\" + re.escape(candidate) + r"\b", text_upper):
        score += 1

    # Generic family IDs (e.g., CHS8600) can repeat across many sheets.
    if re.match(r"^[A-Z]{2,4}\d{4,6}$", candidate):
        score -= 2

    return score


def extract_part_number_diagnostics(text):
    text = (text or "").replace("\n", " ")
    text_upper = text.upper()
    candidates = _capture_candidates_with_sources(text_upper)
    if not candidates:
        return {
            "part_number": None,
            "confidence": 0.0,
            "alternatives": [],
            "sources": [],
            "scores": {},
        }

    scored = []
    has_label_signal = any(
        any(s in {"deinterleaved_label", "direct_label", "first_after_label"} for s in sources)
        for sources in candidates.values()
    )
    prefer_parts_list_structured = not has_label_signal
    for candidate, sources in candidates.items():
        score = _score_candidate(
            candidate,
            sources,
            text_upper,
            prefer_parts_list_structured=prefer_parts_list_structured,
        )
        scored.append((candidate, score, sorted(sources)))

    def _sort_key(row):
        candidate, score, sources = row
        label_preferred = any(
            s in {"deinterleaved_label", "direct_label", "first_after_label"}
            for s in sources
        )
        return (score, label_preferred, len(candidate))

    scored.sort(key=_sort_key, reverse=True)
    best_candidate, best_score, best_sources = scored[0]
    second_score = scored[1][1] if len(scored) > 1 else -1
    margin = best_score - second_score

    if len(scored) == 1:
        confidence = 0.95 if best_score >= 8 else 0.8
    else:
        confidence = min(0.98, 0.58 + (0.07 * margin) + (0.02 * best_score))
        if margin <= 1:
            confidence = min(confidence, 0.72)

    return {
        "part_number": best_candidate,
        "confidence": round(confidence, 3),
        "alternatives": [c for c, _, _ in scored[1:4]],
        "sources": best_sources,
        "scores": {c: s for c, s, _ in scored},
    }


def _has_interleaved_label(compact_tail, label):
    j = 0
    for ch in compact_tail:
        if j < len(label) and ch == label[j]:
            j += 1
    return j >= len(label) - 1


def extract_part_number(text):
    return extract_part_number_diagnostics(text)["part_number"]


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
    manual_review_count = 0
    details = []

    with pdfplumber.open(str(input_pdf)) as pdf:
        for i, page in enumerate(pdf.pages):
            page_num = i + 1
            text = page.extract_text() or ""
            diagnostics = extract_part_number_diagnostics(text)
            part_number = diagnostics["part_number"]
            confidence = diagnostics["confidence"]
            alternatives = diagnostics["alternatives"]
            sources = diagnostics["sources"]

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
                        "confidence": confidence,
                        "sources": sources,
                        "alternatives": alternatives,
                        "output_path": None,
                        "message": message,
                    }
                )
                continue

            if confidence < CONFIDENCE_THRESHOLD:
                review_dir = output_dir / MANUAL_REVIEW_FOLDER_NAME
                review_dir.mkdir(parents=True, exist_ok=True)
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                safe_guess = re.sub(r'[\\/*?:"<>|]', "_", part_number)[:40]
                review_path = review_dir / f"PAGE_{page_num:04d}_{safe_guess}.pdf"
                with open(review_path, "wb") as f:
                    writer.write(f)
                message = (
                    f"Manual review page {page_num}: low confidence {confidence:.3f} "
                    f"for {part_number} -> {review_path}"
                )
                print(message)
                manual_review_count += 1
                details.append(
                    {
                        "page": page_num,
                        "part_number": part_number,
                        "status": "manual_review",
                        "reason": "low_confidence",
                        "confidence": confidence,
                        "sources": sources,
                        "alternatives": alternatives,
                        "output_path": str(review_path),
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
                        "confidence": confidence,
                        "sources": sources,
                        "alternatives": alternatives,
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
                            "confidence": confidence,
                            "sources": sources,
                            "alternatives": alternatives,
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
                        "confidence": confidence,
                        "sources": sources,
                        "alternatives": alternatives,
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
                    "confidence": confidence,
                    "sources": sources,
                    "alternatives": alternatives,
                    "output_path": str(output_path),
                    "message": message,
                }
            )

    not_imported_count = (
        duplicate_in_run_count
        + already_exists_count
        + no_part_number_count
        + purchasing_folder_missing_count
        + manual_review_count
    )

    print(f"Generated {imported_count} unique PDF(s).")
    print(f"Skipped duplicates in input: {duplicate_in_run_count}")
    print(f"Skipped existing files: {already_exists_count}")
    print(f"Skipped pages with no part number: {no_part_number_count}")
    print(f"Skipped purchasing parts (folder not selected): {purchasing_folder_missing_count}")
    print(f"Manual review required: {manual_review_count}")
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
        "manual_review_count": manual_review_count,
        "details": details,
    }


def main():
    process_pdf(INPUT_PDF, OUTPUT_FOLDER, PURCHASING_FOLDER)


if __name__ == "__main__":
    main()
