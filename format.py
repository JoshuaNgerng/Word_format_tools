import sys
import re

from docx.table import Table

from doc_manager import DocManager, SpecialStyle
from detect_missing_headings import HeadingDetection

def infer_and_fix_headings(src_doc_file: str, template_doc_file: str):
    src_doc = DocManager(src_doc_file)
    dst_doc = DocManager(template_doc_file)
    detect = HeadingDetection()

    prev_lv = -1
    # appendix = False
    for block in src_doc.iter_blocks():
        current_lvl = prev_lv
        if isinstance(block, Table):
            dst_doc.copy_table(block)
            continue
        text = block.text.strip()
        if not text:
            continue # ignore empty paragraph in this stage
        
        heading = src_doc.detect_heading(block)
        if not heading:
            lv, to_remove = detect.infer_heading_from_text(text)
            heading = lv
            dst_doc.remove_prefix_from_paragraph(block, to_remove)

        # print(f"{appendix=} {heading=}, {text}")
        if heading is None:
            heading = prev_lv + 1 if prev_lv > 0 else None
        elif isinstance(heading, int):
            current_lvl = heading
        # elif heading == SpecialStyle.Appendix:
            # appendix = True

        block = dst_doc.copy_paragraph(block)

        if isinstance(heading, int):
            dst_doc.apply_heading(block, heading)
        elif isinstance(heading, SpecialStyle):
            dst_doc.apply_subpoint(block, heading)

        # dst_doc.copy_paragraph(block)
        if current_lvl != prev_lv:
            dst_doc.add_empty_paragraph()
        prev_lv = current_lvl

    return dst_doc
    

def cleanup_subpoints(src_doc_file: str | DocManager):
    def clean_paragraph_runs(paragraph):
        """
        Cleans a paragraph by:
        - Capitalizing the first letter of the first run
        - Removing trailing redundant words/punctuation from the last run

        Preserves paragraph structure, numbering, and run formatting.
        """

        runs = paragraph.runs

        if not runs:
            return

        # Clean first run: capitalize first letter if needed
        first_run = runs[0]
        if first_run.text:
            text = first_run.text
            if text[0].islower():
                first_run.text = text[0].upper() + text[1:]

        # Clean last run: remove trailing ; . and optional and/or
        last_run = runs[-1]
        if last_run.text:
            last_run.text = re.sub(
                r'\s*[;,.]\s*(?:and|or)?\s*$',
                '',
                last_run.text
            ).rstrip()

    src_doc = DocManager(src_doc_file) if isinstance(src_doc_file, str) else src_doc_file

    prev_style = False
    for block in src_doc.iter_blocks():
        if isinstance(block, Table):
            continue
        current_style = False
        numbering = src_doc.get_numbering_info(p=block)
        if numbering:
            if numbering.is_numeric:
                prev_style = False
                continue
            if numbering.is_alpha:
                # print(f'detect {block.style} ({prev_style=})| {block.text}')
                if not prev_style:
                    src_doc.reset_subpoint_numbering(
                        block, style_name=None,
                        abstract_id=numbering.abstract_num_id
                    )
                current_style = True
            clean_paragraph_runs(block)
        else:
            name = block.style.name if block.style else None
            if not name: continue
            if name not in src_doc.SUBPOINT_STYLES.values():
                prev_style = False
                continue
            if name == src_doc.SUBPOINT_STYLES[SpecialStyle.Alphabet]:
                if not prev_style:
                    src_doc.reset_subpoint_numbering(
                        block, style_name=name
                    )
                current_style = True
            clean_paragraph_runs(block)

        prev_style = current_style

    return src_doc

def normalize_paragraph_spacing(src_doc_file: str | DocManager, template_doc_file: str):
    current_level = -1
    src_doc = DocManager(src_doc_file) if isinstance(src_doc_file, str) else src_doc_file
    dst_doc = DocManager(template_doc_file)
    after_table = False
    for block in src_doc.iter_blocks():
        if isinstance(block, Table):
            dst_doc.copy_table(block)
            dst_doc.add_empty_paragraph(spacing_after=8)
            after_table = True
            continue
        p = block
        if not p.text.strip():
            # print('skip empty text paragraph')
            continue
        style = p.style.name if p.style else None
        style = style if style else ''
        explicit_level = src_doc.STYLES_LEVEL_MAPPING.get(style, None)
        next_level = (
            explicit_level
            if explicit_level is not None
            else current_level
        )
        if not after_table and explicit_level and current_level >= 0:
            if current_level == next_level:
                if current_level != 2:
                    dst_doc.add_empty_paragraph(src_doc.FALLBACK_STYLE)
            elif current_level > next_level:
                dst_doc.add_empty_paragraph(spacing_after=8)
        elif not after_table and current_level < 0:
            dst_doc.add_empty_paragraph()
        elif after_table:
            after_table = False

        dst_doc.sanitize_whitespace(p)

        if next_level == 2 and current_level != 2:
            dst_doc.apply_bold(p)

        dst_doc.copy_paragraph(p)
        current_level = next_level

    return dst_doc


if __name__ == '__main__':
    doc = infer_and_fix_headings(sys.argv[1], "template.docx")
    doc = cleanup_subpoints(doc)
    doc = normalize_paragraph_spacing(doc, "template.docx")
    doc.save("test.docx")
    # infer_and_fix_headings(sys.argv[1], sys.argv[2])