from dataclasses import dataclass
import re
import argparse
from pathlib import Path

from typing import Any, ClassVar, Generator, Iterator, Protocol
from copy import deepcopy
from docx import Document
from docx.shared import Inches , Pt
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from format_v2 import DocumentFormatter
from number_manager import HeadingNumberingManager

@dataclass
class EmptyParagraph:
    style: str

class NormalizeParagraphs:
    STYLE_MAPPING = {
        '1HeadingStyle': 0,
        '2HeadingStyle': 1,
        '2HeadingLong': 1,
        '3HeadingStyle': 2,
        '4HeadingStyle': 3                            ,
    }
    SPECIAL_STYLES = ["AlphabetListing", "BulletList"]
    EMPTY_STYLE = 'NoNumber'
    DEFAULT = 'Normal'
    MANAGER = DocumentFormatter(HeadingNumberingManager)
    TEMPLATE_PATH = 'template.docx'

    def format(self, input_path: str, out_path: str):
        doc = Document(input_path)
        out = Document(self.TEMPLATE_PATH)
        for block in self._normalize_paragraphs(doc):
            if isinstance(block, Table):
                self.MANAGER.copy_table(block, out)
                continue
            if isinstance(block, Paragraph):
                self.MANAGER.copy_paragraph(block, out)
                continue
            self.MANAGER.add_empty_paragraph(out, block.style)

        out.save(out_path)

    def _normalize_paragraphs(self, doc: _Document) -> Iterator[Paragraph | EmptyParagraph | Table]:
        current_level = -1
        after_table = False
        nm = self.MANAGER.nm(doc)

        for block in nm.iter_blocks(doc):
            print(f'current level {current_level}')
            if isinstance(block, Table):
                # Table separates content visually
                print('detect empty cont')
                yield EmptyParagraph(self.EMPTY_STYLE)
                print('yield block')
                yield block

                # Keep current_level!
                after_table = True
                continue

            p = block

            # if after_table:
            #     print('detect empty end')
            #     yield EmptyParagraph(self.DEFAULT)

            if not p.text:
                print('skip empty text paragraph')
                continue

            style = p.style.name if p.style else None
            style = style if style else ''

            if style in self.SPECIAL_STYLES:
                yield p
                continue

            explicit_level = self.STYLE_MAPPING.get(style)
            print(f'{explicit_level=} {str(style)}')
            # Unknown style inherits current context
            next_level = (
                explicit_level
                if explicit_level is not None
                else current_level
            )
            print(f'{current_level=}, {next_level=}, {current_level==next_level}')
            if not after_table and current_level >= 0:
                if current_level == next_level:
                    print('detect empty cont')
                    yield EmptyParagraph(self.EMPTY_STYLE)

                elif current_level > next_level:
                    # e.g. level2 -> level3
                    print('detect empty end')
                    yield EmptyParagraph(self.DEFAULT)
            else:
                yield EmptyParagraph(self.DEFAULT)
                # current_level > next_level:
                # moving upward, no separator

            after_table = False
            print(f'yield |{p.text}|')
            yield p
            current_level = next_level

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
            f"{input_path.stem}_normalize_spacing.docx"
        )

    doc_manager = NormalizeParagraphs()
    doc_manager.format(str(input_path), str(output_path))
