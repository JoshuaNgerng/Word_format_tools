from typing import ClassVar

from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.document import Document as Doc
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
# from docx.text.paragraph import Paragraph
from dataclasses import dataclass

class Picture:
    def __init__(self, paragraph):
        self.paragraph = paragraph
        self.part = paragraph.part

    def get_images(self):
        images = []

        for run in self.paragraph.runs:
            drawings = run._r.xpath(".//w:drawing")

            for drawing in drawings:
                blips = drawing.xpath(".//a:blip")

                if not blips:
                    continue

                rId = blips[0].get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                )

                if rId:
                    images.append(
                        self.part.related_parts[rId]
                    )

        return images

@dataclass(slots=True)
class NumberingInfo:
    numId: int
    level: int
    depth: int
    format: str
    is_numeric: bool
    is_alpha: bool
    is_bullet: bool
    pattern: str
    text: str

class HeadingNumberingManager:
    """
    Uses the numbering already defined in CorporateTemplate.docx.

    Levels:
        0 -> 1HeadingStyle
        1 -> 2HeadingStyle
        2 -> 3HeadingStyle
        >=3 -> List Number
    """

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    HEADING_STYLES = [
        "1HeadingStyle",
        "2HeadingStyle",
        "3HeadingStyle",
        "4HeadingStyle",
        "AlphabetListing",
        "BulletList"
    ]

    FALLBACK_STYLE : ClassVar[str] = "NoNumber"

    def __init__(self, doc):
        self.doc = doc
        self.lim = len(self.HEADING_STYLES) - 1
        self.numbering = (
            doc.part
            .numbering_part
            .numbering_definitions
            ._numbering
        )
        self.style_id_map = self._get_style_id_map(doc)
        self.numId = 0
        self.numbering_part = doc.part.numbering_part
        self.numbering = (
            self.numbering_part.numbering_definitions._numbering
        )
        self.NS = { 'w': self.W_NS }
        self.fallback_num_id = None
        try:
            self.fallback_num_id = self._create_fallback_instance()
        except: pass
            

    def apply(self, paragraph: Paragraph, level: int):

        style = self.FALLBACK_STYLE
        if level <= self.lim:
            style = self.HEADING_STYLES[level]
            print(f'inspect apply: {style} | {paragraph.text}')

        # print(f'apply {level}, {paragraph.text}')
        paragraph.style = self.doc.styles[style]

    def get_numbering_info(self, paragraph: Paragraph):
        p = paragraph._p

        numPr = p.find(".//w:numPr", self.NS)
        if numPr is None:
            # print('no numPr')
            return None

        numId = numPr.find("w:numId", self.NS)
        ilvl = numPr.find("w:ilvl", self.NS)

        if numId is None:
            # print("no numId")
            return None

        numId = int(numId.get(qn("w:val")))
        level = int(ilvl.get(qn("w:val"))) if ilvl is not None else 0

        # Find matching <w:num>
        num = self.numbering.find(f'./w:num[@w:numId="{numId}"]', self.NS)
        if num is None:
            # print('no num')
            return None

        try:
            abstractNumId = int(
                num.find("w:abstractNumId", self.NS).get(qn("w:val"))
            )
        except: print('no abstractNumId');return None

        abstract = self.numbering.find(
            f'./w:abstractNum[@w:abstractNumId="{abstractNumId}"]',
            self.NS,
        )

        lvl = abstract.find(f'./w:lvl[@w:ilvl="{level}"]', self.NS)

        numFmt = lvl.find("w:numFmt", self.NS).get(qn("w:val"))
        lvlText = lvl.find("w:lvlText", self.NS).get(qn("w:val"))

        numeric_formats = {
            "decimal",
            "decimalZero",
        }
        alpha_formats = {
            "upperRoman",
            "lowerRoman",
            "upperLetter",
            "lowerLetter",
            "ordinal",
            "cardinalText",
        }

        return NumberingInfo(
            numId=numId, level=level,
            depth=level+1, format=numFmt, 
            is_numeric=(numFmt in numeric_formats), 
            is_alpha=(numFmt in alpha_formats),
            is_bullet=(numFmt == "bullet"),
            pattern=lvlText, text=paragraph.text
        )

    def iter_blocks(self, parent: Doc):
        parent_elm = parent.element.body

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                para = Paragraph(child, parent)

                # if child.xpath(".//w:drawing"):
                #     yield Picture(para)
                # else:
                yield para

            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def identify_paragraph_style_lv(self, paragraph: Paragraph):
        pPr = paragraph._p.pPr
        if pPr is None or pPr.pStyle is None:
            return None
        style_id = pPr.pStyle.get(qn("w:val"))
        if style_id not in self.style_id_map:
            return None
        style = self.style_id_map[style_id]
        val = None
        for idx, s in enumerate(self.HEADING_STYLES):
            if s == style:
                val = idx
                break
        if val: return val
        if style == self.FALLBACK_STYLE: return 10
        return None

    # @classmethod
    # def inspect_pargraph_pr(cls, paragraph: Paragraph):
    #     from docx.oxml.ns import qn

    #     for p in doc.paragraphs:
    #         pPr = p._p.pPr

    #         if pPr is not None and pPr.pStyle is not None:
    #             style_id = pPr.pStyle.get(qn("w:val"))
    #             print(repr(p.text), "->", style_id)

    def reset_numbering(self, paragraph):
        numbering = self.doc.part.numbering_part.numbering_definitions._numbering

        pPr = paragraph._p.get_or_add_pPr()

        # Find current numId from paragraph or style
        current_num_id = None

        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            numId = numPr.find(qn("w:numId"))
            if numId is not None:
                current_num_id = int(numId.get(qn("w:val")))

        # If paragraph has no numPr, get it from style
        if current_num_id is None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))

                styles = self.doc.part.styles.element

                style = styles.find(
                    ".//w:style[@w:styleId='{}']".format(style_id),
                    namespaces=styles.nsmap
                )

                if style is not None:
                    style_numPr = style.find(".//w:numPr", namespaces=style.nsmap)

                    if style_numPr is not None:
                        style_numId = style_numPr.find(qn("w:numId"))
                        if style_numId is not None:
                            current_num_id = int(style_numId.get(qn("w:val")))

        if current_num_id is None:
            raise ValueError("Paragraph is not numbered")

        # Find the abstractNumId used by this numId
        abstract_num_id = None

        for num in numbering.num_lst:
            if num.numId == current_num_id:
                abstract = num.find(qn("w:abstractNumId"))
                abstract_num_id = abstract.get(qn("w:val"))
                break

        if abstract_num_id is None:
            raise ValueError("Cannot find abstract numbering")

        # Generate new numId
        existing_ids = [
            n.numId for n in numbering.num_lst
        ]

        new_num_id = max(existing_ids or [0]) + 1

        # Create new numbering instance
        new_num = OxmlElement("w:num")
        new_num.set(qn("w:numId"), str(new_num_id))

        abstract = OxmlElement("w:abstractNumId")
        abstract.set(qn("w:val"), abstract_num_id)

        new_num.append(abstract)

        # Restart level 0
        lvl_override = OxmlElement("w:lvlOverride")
        lvl_override.set(qn("w:ilvl"), "0")

        start_override = OxmlElement("w:startOverride")
        start_override.set(qn("w:val"), "1")

        lvl_override.append(start_override)
        new_num.append(lvl_override)

        numbering.append(new_num)

        # Replace paragraph numPr
        old_numPr = pPr.find(qn("w:numPr"))

        if old_numPr is not None:
            pPr.remove(old_numPr)

        numPr = OxmlElement("w:numPr")

        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")

        numId = OxmlElement("w:numId")
        numId.set(qn("w:val"), str(new_num_id))

        numPr.append(ilvl)
        numPr.append(numId)

        pPr.append(numPr)

    def get_style_num_id(self, doc, style_id):
        styles = doc.styles.element

        style = styles.find(
            f".//w:style[@w:styleId='{style_id}']",
            namespaces=styles.nsmap
        )

        if style is None:
            return None

        numPr = style.find(".//w:numPr", namespaces=styles.nsmap)

        if numPr is None:
            return None

        numId = numPr.find(qn("w:numId"))

        if numId is None:
            return None

        return int(numId.get(qn("w:val")))

    def set_paragraph_num_id(self, paragraph, num_id):
        pPr = paragraph._p.get_or_add_pPr()

        # find existing numPr
        numPr = pPr.find(qn("w:numPr"))

        if numPr is None:
            numPr = OxmlElement("w:numPr")

            # insert after pStyle
            pStyle = pPr.find(qn("w:pStyle"))

            if pStyle is not None:
                pStyle.addnext(numPr)
            else:
                pPr.append(numPr)

        # remove existing numId only
        old_numId = numPr.find(qn("w:numId"))
        if old_numId is not None:
            numPr.remove(old_numId)

        # add new numId
        new_numId = OxmlElement("w:numId")
        new_numId.set(qn("w:val"), str(num_id))

        numPr.append(new_numId)

    def get_next_num_id(self):
        numbering = (
            self.doc.part.numbering_part
            .numbering_definitions
            ._numbering
        )

        existing = [
            int(num.get(qn("w:numId")))
            for num in numbering.findall(qn("w:num"))
        ]

        return max(existing or [0]) + 1

    def clone_numbering(self, old_num_id):
        numbering = (
            self.doc.part.numbering_part
            .numbering_definitions
            ._numbering
        )

        new_num_id = self.get_next_num_id()

        old_num = None

        for num in numbering.findall(qn("w:num")):
            if int(num.get(qn("w:numId"))) == old_num_id:
                old_num = num
                break

        if old_num is None:
            raise Exception("Cannot find numbering")

        abstract_id = old_num.find(qn("w:abstractNumId"))

        new_num = OxmlElement("w:num")
        new_num.set(qn("w:numId"), str(new_num_id))

        new_abstract = OxmlElement("w:abstractNumId")
        new_abstract.set(
            qn("w:val"),
            abstract_id.get(qn("w:val"))
        )

        new_num.append(new_abstract)

        # restart numbering
        lvl_override = OxmlElement("w:lvlOverride")
        lvl_override.set(qn("w:ilvl"), "0")

        start_override = OxmlElement("w:startOverride")
        start_override.set(qn("w:val"), "1")

        lvl_override.append(start_override)
        new_num.append(lvl_override)

        numbering.append(new_num)

        return new_num_id


    def _apply_fallback_numbering(self, paragraph):

        pPr = paragraph._p.get_or_add_pPr()

        existing = pPr.find(
            qn("w:numPr")
        )

        if existing is not None:
            pPr.remove(existing)

        numPr = parse_xml(
            f"""
            <w:numPr xmlns:w="{self.W_NS}">
                <w:ilvl w:val="0"/>
                <w:numId w:val="{self.fallback_num_id}"/>
            </w:numPr>
            """
        )

        pPr.append(numPr)

    def _create_fallback_instance(self):

        style_num_id = self._get_style_num_id(
            self.doc,
            self.FALLBACK_STYLE
        )

        if style_num_id is None:
            raise ValueError(
                f"Style {self.FALLBACK_STYLE} has no numbering"
            )

        abstract_num_id = self._get_abstract_num_id(
            style_num_id
        )

        new_num_id = self._next_num_id()

        num = parse_xml(
            f"""
            <w:num
                xmlns:w="{self.W_NS}"
                w:numId="{new_num_id}">
                <w:abstractNumId
                    w:val="{abstract_num_id}"/>
            </w:num>
            """
        )

        self.numbering.append(num)

        return new_num_id

    def _get_style_num_id(self, doc, style_name):

        style = doc.styles[style_name]

        numPr = style._element.pPr.numPr

        if numPr is None:
            return None

        numId = numPr.find(qn("w:numId"))

        if numId is None:
            return None

        return int(numId.get(qn("w:val")))

    def _get_abstract_num_id(self, num_id):

        for num in self.numbering.findall(qn("w:num")):

            current = int(
                num.get(qn("w:numId"))
            )

            if current == num_id:

                abstract = num.find(
                    qn("w:abstractNumId")
                )

                return int(
                    abstract.get(qn("w:val"))
                )

        return None

    def _next_num_id(self):

        max_id = 0

        for node in self.numbering.findall(
            qn("w:num")
        ):

            value = int(
                node.get(qn("w:numId"))
            )

            max_id = max(
                max_id,
                value
            )

        return max_id + 1

    def _get_style_id_map(self, doc: Doc) -> dict[str, str]:
        res = {}
        for style in doc.styles:
            name = style.name
            if (
                name in self.HEADING_STYLES or
                name == self.FALLBACK_STYLE
            ):
                res[style.style_id] = name
        return res




if __name__ == '__main__':
    from docx import Document

    doc = Document("/Users/LENOVO/Desktop/work/doc_formater/check_format/template.docx")
    mgr = HeadingNumberingManager(doc)

    id = mgr.get_style_num_id(doc, 48)
    
    p = doc.add_paragraph("PURPOSE")
    mgr.apply(p, 0)

    p = doc.add_paragraph("Purpose")
    mgr.apply(p, 1)

    p = doc.add_paragraph("Scope")
    mgr.apply(p, 2)

    p = doc.add_paragraph("Implementation detail")
    mgr.apply(p, 3)

    p = doc.add_paragraph("More detail")
    mgr.apply(p, 4)
    doc.save("check.docx")