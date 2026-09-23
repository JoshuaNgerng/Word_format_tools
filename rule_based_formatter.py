# import shutil
import re
from number_manager import HeadingNumberingManager
from parser_docx import parse_document, ParagraphNode, TableNode, DocumentNode
from copy import deepcopy
from dataclasses import dataclass, field
from docx import Document
from docx.document import Document as Doc
from docx.oxml.ns import qn
from docx.shared import Pt, Inches
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH

class RuleBasedFormatter:

    @dataclass
    class NumberContext:
        previous_explicit: ParagraphNode | None = None
        pending: list[ParagraphNode] = field(default_factory=list)


    def __init__(self, numbering_manager_cls):
        self.numbering_manager_cls = numbering_manager_cls

    def format(self, input_path: str, output_path: str):

        src = Document(input_path)

        nodes = parse_document(src)
        self._assign_levels(nodes)
        out = Document()

        numbering = self.numbering_manager_cls(out)

        self._build_output(nodes, out, numbering)

        out.save(output_path)

    def _assign_levels(self, nodes):

        ctx = RuleBasedFormatter.NumberContext()

        # start = self._find_body_start(nodes)

        for node in nodes:

            if not isinstance(node, ParagraphNode):
                continue

            # Existing numbered heading
            if node.explicit_number:

                self._flush_pending(ctx, node)

                node.target_level = self._explicit_level(node)
                ctx.previous_explicit = node

                continue

            # Ignore bullets etc.
            if self._is_exception(node):
                continue

            ctx.pending.append(node)

        self._flush_pending(ctx, None)

    # ---------- helpers ----------
    def _flush_pending(
        self,
        ctx: NumberContext,
        next_explicit: ParagraphNode | None,
    ):

        if not ctx.pending:
            return

        if self._has_conflict(ctx.previous_explicit, next_explicit):

            for p in ctx.pending:
                p.target_level = None

        else:

            level = (self._explicit_level(ctx.previous_explicit) or 0) + 1

            for p in ctx.pending:
                p.target_level = level

        ctx.pending.clear()

    def _find_body_start(self, nodes) -> int:
        """
        Returns the index after the TOC.
        If no TOC exists, start from beginning.
        """

        for i, node in enumerate(nodes):
            if not isinstance(node, ParagraphNode):
                continue

            if node.kind == "toc":
                return i + 1

        return 0

    def _is_exception(self, node: ParagraphNode) -> bool:
        """
        Paragraphs that should never receive generated numbering.
        """

        if node.is_list:
            return True

        if node.kind in ("bullet", "alpha"):
            return True

        return False

    def _explicit_level(self, node: ParagraphNode | None) -> int | None:
        """
        Returns numbering level if paragraph already has an explicit number.

        1.0      -> 1
        1.2      -> 2
        1.2.3    -> 3
        """
        if not node or not node.explicit_number:
            return None

        parts = node.explicit_number.split(".")

        # Word ilvl:
        #
        # 1.0       -> 0
        # 1.1       -> 1
        # 1.1.1     -> 2
        #
        return len(parts) - 2 if parts[-1] == "0" else len(parts) - 1

    def _find_previous_explicit(
        self,
        nodes,
        idx,
    ) -> ParagraphNode | None:

        for i in range(idx - 1, -1, -1):

            node = nodes[i]

            if not isinstance(node, ParagraphNode):
                continue

            if node.explicit_number:
                return node

        return None

    def _find_next_explicit(
        self,
        nodes,
        idx,
    ) -> ParagraphNode | None:

        for i in range(idx + 1, len(nodes)):

            node = nodes[i]

            if not isinstance(node, ParagraphNode):
                continue

            if node.explicit_number:
                return node

        return None


    def _infer_missing_level(
        self,
        prev: ParagraphNode,
        nxt: ParagraphNode,
    ) -> int | None:

        if self._has_conflict(prev, nxt):
            return None

        return (self._explicit_level(prev) or 0) + 1

    def _has_conflict(
        self,
        prev: ParagraphNode | None,
        nxt: ParagraphNode | None,
    ) -> bool:

        if prev is None or nxt is None:
            return True

        prev_level = self._explicit_level(prev)
        next_level = self._explicit_level(nxt)

        if prev_level != next_level:
            return True

        return False

    def _build_output(
        self,
        nodes: DocumentNode,
        out: Doc,
        numbering: NumberingManager,
    ):

        for node in nodes:

            if isinstance(node, ParagraphNode):
                self._copy_paragraph(
                    out,
                    node,
                    numbering,
                )

            elif isinstance(node, TableNode):
                self._copy_table(out, node)

    def _copy_table(
        self,
        out: Doc,
        table_node: TableNode,
    ):

        body = out._element.body

        body.append(
            deepcopy(table_node.table)
        )

    def _copy_paragraph(
        self,
        out: Doc,
        node: ParagraphNode,
        numbering: NumberingManager,
    ):

        # Clone original paragraph
        new_p = deepcopy(node.paragraph._p)

        out._element.body.append(new_p)

        dst = Paragraph(
            new_p,
            out
        )

        #
        # Remove explicit numbering text
        #
        self._remove_explicit_number(
            dst,
            node.explicit_number,
        )

        #
        # Keep original formatting, only normalize font
        #
        # for i, run in enumerate(dst.runs):
        #     print(
        #         i,
        #         repr(run.text),
        #         run.bold,
        #         run.font.color.rgb,
        #     )

        self._normalize_runs(dst)

        #
        # Apply numbering / indent
        #
        print(f'{node.target_level}: {node.text}')
        if node.target_level is not None:
            
            numbering.apply(
                dst,
                node.target_level,
            )

        else:

            fmt = dst.paragraph_format
            fmt.left_indent = Inches(0.5)
            fmt.first_line_indent = Inches(0)

    def _normalize_runs(
        self,
        paragraph: Paragraph,
    ):

        for run in paragraph.runs:

            font = run.font

            font.name = "Arial"
            font.size = Pt(10)

            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()

            rFonts.set(qn("w:ascii"), "Arial")
            rFonts.set(qn("w:hAnsi"), "Arial")
            rFonts.set(qn("w:cs"), "Arial")
            rFonts.set(qn("w:eastAsia"), "Arial")

    def _remove_explicit_number(
        self,
        paragraph,
        explicit_number: str | None,
    ):

        if explicit_number is None:
            return

        pattern = re.compile(
            rf"^\s*{re.escape(explicit_number)}\s*"
        )

        for run in paragraph.runs:

            new_text = pattern.sub(
                "",
                run.text,
                count=1,
            )

            if new_text != run.text:

                run.text = new_text
                return

    def _apply_style(self, paragraph):
        self._apply_font(paragraph)
        self._apply_normal_format(paragraph)

    def _apply_normal_format(
        self,
        paragraph
    ):

        fmt = paragraph.paragraph_format

        fmt.left_indent = Inches(0.5)

        fmt.first_line_indent = Inches(0)

        paragraph.alignment = (
            WD_ALIGN_PARAGRAPH.LEFT
        )

    def _apply_font(
        self,
        paragraph
    ):

        for run in paragraph.runs:

            run.font.name = "Arial"
            run.font.size = Pt(10)

    def _level_from_number(self, num: str) -> int:
        return num.count(".")
