# ****************************************************************************************************
# Script Name    : pack_pdfs_with_toc.py
#
# Purpose        : Merges all PDFs in a folder, adds bookmarks, and inserts a clickable Table of Contents (TOC).
#
# Author         : Manivannan Mathialagan
# Created On     : 18-May-2025
#
# Example Usage
#   python pack_pdfs_with_toc.py "input_folder_path" "output_file_name.pdf" Y
#
# Notes
#   - Requires Python 3 and PyMuPDF installed (`pip install pymupdf`)
# ****************************************************************************************************

import os
import sys
import re
import fitz  # PyMuPDF

def parse_sort_key(title):
    match = re.search(r'(\d+(\.\d+)*)', title)
    return [int(x) for x in match.group(1).split('.')] if match else [9999]

def extract_title_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = doc[0].get_text("text")
        doc.close()
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        id_title, desc_lines = "", []
        for line in lines[:20]:
            if not id_title:
                match = re.search(r'(Listing|Table|Figure)\s+\d+(\.\d+)*', line)
                if match:
                    id_title = match.group().strip()
                    remainder = line[len(match.group()):].strip(" :–-")
                    if remainder:
                        desc_lines.append(remainder)
            elif len(desc_lines) < 2:
                desc_lines.append(line)
        if id_title:
            final = f"{id_title}: {' - '.join(desc_lines)}"
        elif lines:
            final = lines[0]
        else:
            final = os.path.basename(pdf_path)
        return final
    except Exception as e:
        print(f"[Fallback] Could not extract title from {pdf_path}: {e}")
        return os.path.basename(pdf_path)

def merge_pdfs_with_bookmarks(entries, output_path):
    merged = fitz.open()
    toc = []
    for entry in entries:
        sub = fitz.open(entry['pdf'])
        start = len(merged)
        merged.insert_pdf(sub)
        toc.append([1, entry['title'], start + 1])
        print(f"[Bookmark] {entry['title']} Page {start + 1}")
    merged.set_toc(toc)
    merged.save(output_path)
    print(f"[Saved] Merged PDF: {output_path}")

def wrap_toc_title(title, font_size, max_width):
    words = title.split()
    lines = []
    current_line = ""
    for word in words:
        if current_line == "":
            current_line = word
            continue
        test_line = current_line + " " + word
        if fitz.get_text_length(test_line, fontsize=font_size) <= max_width or len(current_line.split()) == 1:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return [line for line in lines if line.strip()]

def add_toc_and_links(bookmarked_pdf, final_pdf, font_size=8, include_toc_title=False):
    def extract_toc_entries(doc):
        return [(lvl, title.strip(), page_num - 1) for lvl, title, page_num in doc.get_toc(simple=True)]

    def paginate_wrapped_entries(toc_entries, font_size, toc_text_max_width, page_height, include_toc_title):
        y_spacing = font_size * 1.5
        lines_per_page = int((page_height - 100) // y_spacing)
        if include_toc_title:
            lines_per_page -= 2
        paginated_entries, current_page_entries, current_line_count = [], [], 0
        for entry in toc_entries:
            level, title, target_page = entry
            indent = 20 * (level - 1)
            entry_max_width = toc_text_max_width - indent
            wrapped_lines = wrap_toc_title(title, font_size, entry_max_width)
            line_count = len(wrapped_lines)
            if current_line_count + line_count > lines_per_page:
                paginated_entries.append(current_page_entries)
                current_page_entries, current_line_count = [], 0
            current_page_entries.append((entry, wrapped_lines))
            current_line_count += line_count
        if current_page_entries:
            paginated_entries.append(current_page_entries)
        return paginated_entries

    def generate_toc_pages(paginated_entries, font_size, page_width, page_height, toc_page_count, toc_text_max_width, page_number_x, page_number_width, gap, right_margin, include_toc_title):
        toc_doc = fitz.open()
        link_targets = []
        left_margin, top_margin = 50, 50
        y_spacing = font_size * 1.5
        for page_index, entries in enumerate(paginated_entries):
            page = toc_doc.new_page(width=page_width, height=page_height)
            y = top_margin
            if include_toc_title and page_index == 0:
                toc_title = "Table of Contents"
                title_width = fitz.get_text_length(toc_title, fontsize=font_size + 2)
                page.insert_text(
                    (page_width / 2 - title_width / 2, y),
                    toc_title,
                    fontsize=font_size + 2,
                    fontname="Helvetica",
                    color=(0, 0, 0)
                )
                y += font_size * 2
            for (level, title, target_page), wrapped_lines in entries:
                indent = 20 * (level - 1)
                x = left_margin + indent
                page_number_str = str(target_page + toc_page_count + 1)
                page_number_width_local = fitz.get_text_length(page_number_str, fontsize=font_size)
                first_line_y = y
                for i, line in enumerate(wrapped_lines):
                    line_width = fitz.get_text_length(line, fontsize=font_size)
                    if i == len(wrapped_lines) - 1:
                        page.insert_text((x, y), line, fontsize=font_size)
                        dots_start_x = x + line_width + 2
                        if dots_start_x < page_number_x - page_number_width_local - gap:
                            dots_width = page_number_x - page_number_width_local - gap - dots_start_x
                            dot_char_width = fitz.get_text_length('.', fontsize=font_size)
                            dot_count = int(dots_width // dot_char_width)
                            dots_str = '.' * dot_count
                            page.insert_text((dots_start_x, y), dots_str, fontsize=font_size)
                        page.insert_text(
                            (page_number_x - page_number_width_local, y),
                            page_number_str, fontsize=font_size
                        )
                    else:
                        page.insert_text((x, y), line, fontsize=font_size)
                    y += y_spacing
                rect = fitz.Rect(x, first_line_y - font_size, page_number_x, y)
                link_targets.append((page_index, rect, target_page))
        return toc_doc, link_targets

    base = fitz.open(bookmarked_pdf)
    original_toc = extract_toc_entries(base)
    original_bookmarks = base.get_toc()

    width, height = fitz.paper_size("a4")
    left_margin, right_margin = 50, 60

    max_page_number = 99999
    page_number_sample = str(max_page_number)
    page_number_width = fitz.get_text_length(page_number_sample, fontsize=font_size)
    page_number_x = width - right_margin
    gap = 8

    toc_text_max_width = page_number_x - left_margin - page_number_width - gap

    print(f"TOC: width={width:.1f}, toc_text_max_width={toc_text_max_width:.1f}, page_number_x={page_number_x:.1f}, font_size={font_size}")

    paginated = paginate_wrapped_entries(original_toc, font_size, toc_text_max_width, height, include_toc_title)
    toc_doc, link_targets = generate_toc_pages(
        paginated, font_size, width, height, len(paginated),
        toc_text_max_width, page_number_x, page_number_width, gap, right_margin, include_toc_title
    )
    toc_page_count = len(paginated)

    final = fitz.open()
    final.insert_pdf(toc_doc)
    final.insert_pdf(base)

    def add_toc_hyperlinks(doc, link_targets, toc_page_count):
        for toc_page_index, rect, target_page in link_targets:
            doc[toc_page_index].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": rect,
                "page": target_page + toc_page_count
            })
    add_toc_hyperlinks(final, link_targets, toc_page_count)

    def shift_bookmark_pages(bookmarks, offset):
        return [[lvl, title, page + offset] for lvl, title, page in bookmarks if len([lvl, title, page]) >= 3]
    def add_existing_bookmarks(doc, bookmarks, offset):
        doc.set_toc(shift_bookmark_pages(bookmarks, offset))
    add_existing_bookmarks(final, original_bookmarks, toc_page_count)

    final.save(final_pdf)
    print(f"[Success] Final PDF with TOC created: {final_pdf}")
    final.close()
    base.close()

    # Optionally delete the intermediate file (merged without TOC)
    if os.path.exists(bookmarked_pdf) and bookmarked_pdf != final_pdf:
        try:
            os.remove(bookmarked_pdf)
        except Exception as e:
            print(f"[Warning] Couldn't delete intermediate: {e}")

def main(folder, output_pdf, include_toc_title=False):
    folder = os.path.abspath(folder)
    pdf_files = sorted(f for f in os.listdir(folder) if f.lower().endswith('.pdf') and not f.startswith('~$'))

    entries = []
    for pdf in pdf_files:
        pdf_path = os.path.join(folder, pdf)
        title = extract_title_from_pdf(pdf_path)
        entries.append({'pdf': pdf_path, 'title': title})

    entries = sorted(entries, key=lambda x: parse_sort_key(x['title']))
    merged_pdf = os.path.join(folder, "__tmp_merge__.pdf")

    merge_pdfs_with_bookmarks(entries, merged_pdf)
    add_toc_and_links(merged_pdf, output_pdf, font_size=8, include_toc_title=include_toc_title)

if __name__ == "__main__":
    if len(sys.argv) < 3 or len(sys.argv) > 4:
        print("Usage: python pack_pdfs_with_toc.py <folder_path> <output_pdf> [toc_title: Y/N]")
        sys.exit(1)

    path = sys.argv[1]
    output_pdf = sys.argv[2]
    include_toc_title = sys.argv[3].strip().upper() == "Y" if len(sys.argv) > 3 else False

    if not os.path.isdir(path):
        print(f"Error: Folder not found - {path}")
        sys.exit(1)

    main(path, output_pdf, include_toc_title)
