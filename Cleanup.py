import json

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import re
from pathlib import Path
import argparse


class IterList:
    def get_list_info(self, paragraph):
        """
        Returns (numId, ilvl) if paragraph is part of a list.
        Otherwise returns (None, None).
        """
        pPr = paragraph._p.pPr
        if pPr is None:
            return None, None

        numPr = pPr.numPr
        if numPr is None:
            return None, None

        numId = numPr.numId
        ilvl = numPr.ilvl

        return (
            int(numId.val) if numId is not None else None,
            int(ilvl.val) if ilvl is not None else 0,
        )

    def build_numbering_lookup(self, doc):
        """
        Build lookup:
            numId -> (format, level)

        format examples:
            bullet
            decimal
            lowerLetter
            upperLetter
            lowerRoman
        """
        lookup = {}

        numbering = doc.part.numbering_part.numbering_definitions._numbering

        # map abstractNumId -> level -> format
        abstract_formats = {}

        for abstract in numbering.findall(qn("w:abstractNum")):
            abstract_id = int(abstract.get(qn("w:abstractNumId")))
            levels = {}

            for lvl in abstract.findall(qn("w:lvl")):
                ilvl = int(lvl.get(qn("w:ilvl")))
                numFmt = lvl.find(qn("w:numFmt"))
                if numFmt is not None:
                    levels[ilvl] = numFmt.get(qn("w:val"))

            abstract_formats[abstract_id] = levels

        # map numId -> abstractNumId
        for num in numbering.findall(qn("w:num")):
            numId = int(num.get(qn("w:numId")))
            abstractNumId = num.find(qn("w:abstractNumId"))
            if abstractNumId is None:
                continue

            abstract_id = abstractNumId.get(qn("w:val"))
            try:
                abstract_id = int(abstract_id)
            except:
                continue

            lookup[numId] = abstract_formats.get(abstract_id, {})
            with open('debug-numbering-lookup', 'w') as f:
                json.dump(lookup, f, indent=2)
        return lookup

def clean_end(text):
    text = re.sub(
        r'\s*[;.]\s*(?:and|or)?\s*$', '', text
    ).strip()
    if not text: return text
    if text[0].islower():
        text = text[0].upper() + text[1:]
    if len(text) > 2 and (text[-1] == ';' or text[-1] == ',' or text[-1] == '.'):
        return text[:-1]
    return text

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

def get_abstract_id(doc, style_name):
    # extract num id from style name
    style = doc.styles[style_name]
    numPr = style._element.pPr.numPr
    numId = numPr.find(qn("w:numId"))
    num_id = int(numId.get(qn("w:val")))

    # extract abstract num id repr style in this doc
    root = doc.part.numbering_part.element
    for num in root.findall(qn("w:num")):
        if int(num.get(qn("w:numId"))) == num_id:
            abstract = num.find(qn("w:abstractNumId"))
            return int(abstract.get(qn("w:val")))

    return None

def create_new_num(numbering_part, abstract_id):
    root = numbering_part.element

    existing_ids = [
        int(num.get(qn("w:numId")))
        for num in root.findall(qn("w:num"))
    ]

    new_num_id = max(existing_ids, default=0) + 1

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_num_id))

    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), str(abstract_id))
    num.append(abstract)

    # Force restart at 1
    lvl_override = OxmlElement("w:lvlOverride")
    lvl_override.set(qn("w:ilvl"), "0")

    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), "1")

    lvl_override.append(start_override)
    num.append(lvl_override)

    root.append(num)

    return new_num_id

def set_paragraph_num_id(paragraph, num_id):
    pPr = paragraph._p.get_or_add_pPr()

    numPr = pPr.find(qn("w:numPr"))

    if numPr is None:
        numPr = OxmlElement("w:numPr")
        pPr.append(numPr)

    numId = numPr.find(qn("w:numId"))

    if numId is None:
        numId = OxmlElement("w:numId")
        numPr.append(numId)

    numId.set(qn("w:val"), str(num_id))

def main(input_file_path, output_file_path, styles = [], restart_style = None):
    doc = Document(input_file_path)

    checker = IterList()
    lookup = checker.build_numbering_lookup(doc)

    prev_style = None
    abstract_id = None
    if restart_style:
        abstract_id = get_abstract_id(doc, restart_style)

    for p in doc.paragraphs:
        numId, level = checker.get_list_info(p)

        should_clean = False

        if numId is not None:
            fmt = lookup.get(numId, {}).get(level)
            if fmt in {"bullet", "lowerLetter", "upperLetter"}:
                should_clean = True

        if p.style and p.style.name in styles:
            should_clean = True

        if should_clean:
            # print(f'before {p.text}\n{p._p.xml}')
            new_text = clean_end(p.text)
            # print(f'|{p.text} -> {new_text}|')
            p.text = new_text
            # print(f'after {p.text}\n{p._p.xml}')
            # clean_paragraph_runs(p)

        if restart_style and p.style and p.style.name == restart_style:
            if prev_style != restart_style:
                print(f'triiger restart numbering {p.text}')
                # Get style's original abstract definition
                new_num_id = create_new_num(
                    doc.part.numbering_part,
                    abstract_id=abstract_id
                )

                set_paragraph_num_id(
                    p,
                    new_num_id
                )
                print(f'debug {abstract_id=}, {new_num_id=}')

        prev_style = p.style.name if p.style else None 


    doc.save(output_file_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Format an input file."
    )

    parser.add_argument(
        "input_file",
        help="Path to the input file."
    )

    parser.add_argument(
        "-o", "--output",
        help="Path to the output file. "
             "Defaults to '<input_stem>_formatted<input_suffix>'."
    )

    args = parser.parse_args()

    input_path = Path(args.input_file)

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_name(
            f"{input_path.stem}_formatted{input_path.suffix}"
        )

    main(str(input_path), str(output_path), ["BulletList", "AlphabetListing"], "AlphabetListing")

def _basic_check(file_name):
    doc = Document(file_name)

    checker = IterList()
    lookup = checker.build_numbering_lookup(doc)

    for p in doc.paragraphs:

        numId, level = checker.get_list_info(p)

        if numId is None:
            print(f"TEXT     : {p.text}")
            continue

        fmt = lookup.get(numId, {}).get(level, "unknown")

        if fmt == "bullet":
            list_type = "BULLET"

        elif fmt in ("lowerLetter", "upperLetter"):
            list_type = "ALPHABET"

        elif fmt == "decimal":
            list_type = "NUMBER"

        elif fmt in ("lowerRoman", "upperRoman"):
            list_type = "ROMAN"

        else:
            list_type = fmt

        print(f"{list_type:<10} Level {level} : {p.text}")