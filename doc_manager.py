from copy import deepcopy
from functools import lru_cache
from enum import StrEnum, auto
import json
from pathlib import Path
from typing import ClassVar, cast

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.document import Document as Doc
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.image import ImagePart

# from docx.text.paragraph import Paragraph
from dataclasses import dataclass
from dataclasses import asdict, is_dataclass

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
    num_id: int
    abstract_num_id: int
    level: int

    format: str
    level_text: str
    start: int

    is_bullet: bool
    is_ordered: bool
    is_numeric: bool
    is_alpha: bool

class SpecialStyle(StrEnum):
    Alphabet = 'a'
    Bullet = '.'
    Appendix = 'APPENDIX'
    AppendixBody = ''


def serializer(obj):
    if is_dataclass(obj):
        return asdict(obj) # type: ignore
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")

class DocManager:
    """
    Uses the numbering already defined in CorporateTemplate.docx.

    Levels:
        0 -> 1HeadingStyle
        1 -> 2HeadingStyle
        2 -> 3HeadingStyle
        >=3 -> List Number
    """

    @dataclass
    class StyleInfo:
        name: str
        style_id: str
        abstract_id: str

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    HEADING_STYLES = {
        1: ["1HeadingStyle", "AppendixHeading"],
        2: ["2HeadingStyle", "2HeadingLong"],
        3: ["3HeadingStyle"],
        4: ["4HeadingStyle"]
    }

    STYLES_LEVEL_MAPPING = {}

    SUBPOINT_STYLES = {
        SpecialStyle.Alphabet: "AlphabetListing",
        SpecialStyle.Bullet: "BulletList",
        SpecialStyle.Appendix: "AppendixHeading",
        SpecialStyle.AppendixBody: "AppendixBody"
    }

    FALLBACK_STYLE : ClassVar[str] = "NoNumber"

    def __init__(self, file_path: str | Path):
        self.NS = { 'w': self.W_NS }
        if not self.STYLES_LEVEL_MAPPING:
            self.STYLES_LEVEL_MAPPING = self.__map_reverse()
        # map foriegn rid to source rid when importing pic from other doc
        self.doc = Document(str(file_path))
        self.style_mapping = self.__get_style_mapping()
        self.numbering_lookup = self.__build_numbering_lookup()
        with open(f'debug-{file_path}.json', 'w') as f:
            json.dump(self.style_mapping, f, indent=2, default=serializer)

    def save(self, file_path: str | Path):
        self.doc.save(str(file_path))

    def apply_heading(self, paragraph: Paragraph, level: int):
        style = self.FALLBACK_STYLE
        try:
            style = self.HEADING_STYLES[level][0]
        except:
            pass

        paragraph.style = style 


    def apply_subpoint(self, paragraph: Paragraph, type_: SpecialStyle | None):
        if not type_ or not type_ in self.SUBPOINT_STYLES:
            type_ = SpecialStyle.Alphabet
        paragraph.style = self.SUBPOINT_STYLES[type_]

    def detect_heading(self, paragraph: Paragraph) -> int | SpecialStyle | None:
        pPr = paragraph._p.pPr
        if pPr is None or pPr.pStyle is None:
            return None
        style_id = pPr.pStyle.get(qn("w:val"))
        style_name = None
        for style in self.style_mapping.values():
            if style.style_id == style_id:
                style_name = style.name
        if style_name is None: return None
        if style_name not in self.STYLES_LEVEL_MAPPING:
            key = next(
                (
                    k for k, v in self.SUBPOINT_STYLES.items() 
                    if v == style_name
                ), None
            )
            return key
        return self.STYLES_LEVEL_MAPPING[style_name]

    def reset_subpoint_numbering(
            self, paragraph: Paragraph, style_name: str | None, 
            abstract_id: int | str | None = None
    ):
        if not abstract_id:
            try:
                abstract_id = (
                    self
                        .style_mapping[style_name] # type: ignore
                        .abstract_id
                )
            except:
                abstract_id = None
        if not abstract_id:
            print('triiger falier?')
            return
        new_num_id = self.__create_new_num(
            self.doc.part.numbering_part,
            abstract_id=abstract_id
        )
        self.__set_paragraph_num_id(
            paragraph,
            new_num_id
        )
        print(f'triiger sucess reset numbering {abstract_id=}, {new_num_id=}')

    def iter_blocks(self):
        parent = self.doc
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

    def copy_paragraph(self, src: Paragraph):
        dst = self.doc.add_paragraph()

        # Cache source rId -> destination rId.
        # This prevents creating multiple relationships when the
        # same image is used more than once.
        rid_map = {}

        for run in src.runs:
            drawing = run._r.find(qn("w:drawing"))

            if drawing is not None:
                self._copy_drawing(
                    src=src,
                    dst=dst,
                    drawing=drawing,
                    rid_map=rid_map,
                )
            else:
                new = dst.add_run(run.text)

                new.bold = run.bold
                new.italic = run.italic
                new.underline = run.underline

                # Preserve colour
                if run.font.color.rgb:
                    new.font.color.rgb = run.font.color.rgb

                # Enforce font
                new.font.name = "Arial"
                new.font.size = Pt(10)

        dst.style = src.style

        return dst

    def copy_table(self, src_table: Table):
        self.doc._body._element._insert_tbl(deepcopy(src_table._tbl))

    def add_empty_paragraph(self, style=None, size=None, spacing_after=None):
        empty = self.doc.add_paragraph()

        # enforce formatting if required
        run = empty.add_run("")
        run.font.name = "Arial"
        run.font.size = Pt(size or 10)

        if style: empty.style = style
        if spacing_after:
            empty.paragraph_format.space_after = Pt(spacing_after)

        return empty

    def sanitize_whitespace(self, p: Paragraph):
        runs = p.runs
        if not runs: return
        first_run = runs[0]
        first_run.text = first_run.text.strip()
        last_run = runs[-1]
        last_run.text = last_run.text.strip()

    def apply_bold(self, p: Paragraph):
        for run in p.runs:
            run.bold = True

    def refresh_numbering(self):
        self.numbering_lookup = self.__build_numbering_lookup()

    def get_numbering_info(self, p: Paragraph):
        if not self.numbering_lookup:
            self.numbering_lookup = self.__build_numbering_lookup()
        numId, ilvl = self._get_list_info(p)
        if numId is None or ilvl is None:
            return None
        try:
            res = self.numbering_lookup[(numId, ilvl)]
        except:
            res = None
        return res

    @classmethod
    def remove_prefix_from_paragraph(cls, p: Paragraph, prefix: str) -> Paragraph:
        if not prefix:
            return p

        runs = p.runs
        full_text = "".join(run.text or "" for run in runs)

        # Ignore whitespace at the start of the paragraph
        start = len(full_text) - len(full_text.lstrip())

        # Only check the prefix at the start.
        # We NEVER search for it elsewhere.
        if not full_text.startswith(prefix, start):
            return p

        remove_start = start
        remove_end = start + len(prefix)

        # Map the character range back onto the individual runs
        position = 0

        for run in runs:
            text = run.text or ""
            run_start = position
            run_end = position + len(text)

            if run_end <= remove_start:
                # This run is entirely before the prefix
                position = run_end
                continue

            if run_start >= remove_end:
                # Prefix has already been removed
                break

            # Calculate which part of this run overlaps the prefix
            local_start = max(remove_start - run_start, 0)
            local_end = min(remove_end - run_start, len(text))

            run.text = (
                text[:local_start] +
                text[local_end:]
            )

            position = run_end

        return p

    def _get_list_info(self, p: Paragraph):
        """
        Returns (numId, ilvl) for the paragraph's effective numbering.

        Checks:
        1. Explicit paragraph numbering
        2. Numbering inherited from the paragraph style

        Returns:
            (numId, ilvl)
            (None, None) if the paragraph is not numbered.
        """

        # ------------------------------------------------------------
        # 1. Explicit paragraph numbering
        # ------------------------------------------------------------
        pPr = p._p.pPr

        if pPr is not None and pPr.numPr is not None:
            numPr = pPr.numPr

            numId = numPr.numId
            ilvl = numPr.ilvl

            if numId is not None:
                return (
                    int(numId.val),
                    int(ilvl.val) if ilvl is not None else 0,
                )

        # ------------------------------------------------------------
        # 2. Numbering inherited from paragraph style
        # ------------------------------------------------------------
        style = p.style

        if style is not None:
            style_pPr = style._element.pPr # type: ignore hidden element

            if style_pPr is not None and style_pPr.numPr is not None:
                numPr = style_pPr.numPr

                numId = numPr.numId
                ilvl = numPr.ilvl

                if numId is not None:
                    return (
                        int(numId.val),
                        int(ilvl.val) if ilvl is not None else 0,
                    )

        return None, None

    def _copy_drawing(
        self,
        src: Paragraph,
        dst: Paragraph,
        drawing,
        rid_map: dict[str, str],
    ):
        """
        Copy a <w:drawing> from src to dst, copying any referenced
        image relationship and rewriting r:embed to the new rId.
        """

        new_run = dst.add_run()

        # Copy the drawing XML.
        new_drawing = deepcopy(drawing)

        assert self.doc.part.package
        # Find image relationship IDs used by this drawing.
        #
        # Normal DrawingML images look like:
        #
        #   <a:blip r:embed="rId7"/>
        #
        # There can also be r:link for externally linked images.
        for element in new_drawing.iter():

            source_rid = element.get(qn("r:embed"))

            if source_rid is None:
                continue

            # Already copied this relationship.
            if source_rid in rid_map:
                destination_rid = rid_map[source_rid]
            else:
                # Relationship belongs to the part containing the
                # source paragraph.
                source_part = src.part

                source_rel = source_part.rels[source_rid]

                if source_rel.reltype != RT.IMAGE:
                    raise ValueError(
                        f"Expected image relationship for {source_rid}, "
                        f"got {source_rel.reltype}"
                    )

                source_image_part = source_rel.target_part

                if not isinstance(source_image_part, ImagePart):
                    raise TypeError(
                        f"Expected ImagePart, got "
                        f"{type(source_image_part).__name__}"
                    )

                # Try to reuse an identical image already in destination.
                destination_image_part = None

                for existing_part in self.doc.part.package.image_parts:
                    if existing_part.sha1 == source_image_part.sha1:
                        destination_image_part = existing_part
                        break

                # Image doesn't exist yet in destination.
                if destination_image_part is None:
                    destination_image_part = (
                        self.doc.part.package.image_parts
                        ._add_image_part(source_image_part.image)
                    )

                # Create relationship from destination document -> image.
                destination_rid = self.doc.part.relate_to(
                    destination_image_part,
                    RT.IMAGE,
                )

                rid_map[source_rid] = destination_rid

            # Replace source rId with destination rId.
            element.set(
                qn("r:embed"),
                destination_rid,
            )

        # Add the fixed drawing XML to the new run.
        new_run._r.append(new_drawing)

    def __create_new_num(self, numbering_part, abstract_id):
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

    def __set_paragraph_num_id(self, paragraph, num_id):
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

    def __get_abstract_id(self, style_name):
        # extract num id from style name
        style = self.doc.styles[style_name]
        numPr = style._element.pPr.numPr # type: ignore
        if numPr is None:
            return None
        numId = numPr.find(qn("w:numId"))
        num_id = int(numId.get(qn("w:val")))

        # extract abstract num id repr style in this doc
        root = self.doc.part.numbering_part.element
        for num in root.findall(qn("w:num")):
            if int(num.get(qn("w:numId"))) == num_id:
                abstract = num.find(qn("w:abstractNumId"))
                return str(abstract.get(qn("w:val")))

        return None

    def __get_style_mapping(self) -> dict[str, StyleInfo]:
        '''
        name -> StyleInfo
        '''
        res = {}
        for style in self.doc.styles:
            name = style.name
            if not name:
                continue
            if (
                name in self.STYLES_LEVEL_MAPPING.keys() or 
                name in self.SUBPOINT_STYLES.values()
            ):
                res[name] = self.StyleInfo(
                    name=name, style_id=style.style_id, 
                    abstract_id=self.__get_abstract_id(name) or ''
                )
        return res

    def __map_reverse(self):
        res : dict[str, int] = {}
        for level, styles in self.HEADING_STYLES.items():
            for style in styles:
                res[style] = level
        return res

    def __build_numbering_lookup(
        self
    ) -> dict[tuple[int, int], NumberingInfo]:
        """
        Build a lookup for the numbering definitions in the document.

        Key:
            (numId, ilvl)

        Example:
            (7, 0) -> NumberingInfo(...)
            (7, 1) -> NumberingInfo(...)

        This allows a single numId to have different definitions at
        different indentation levels.
        """
        doc = self.doc
        try:
            numbering = doc.part.numbering_part.element
        except:
            return {}

        # ------------------------------------------------------------------
        # 1. Build:
        #
        #     abstractNumId -> ilvl -> <w:lvl>
        #
        # ------------------------------------------------------------------
        abstract_levels: dict[int, dict[int, object]] = {}

        for abstract in numbering.findall(qn("w:abstractNum")):
            abstract_id = abstract.get(qn("w:abstractNumId"))
            if abstract_id is None:
                continue

            abstract_id = int(abstract_id)

            levels: dict[int, object] = {}

            for lvl in abstract.findall(qn("w:lvl")):
                ilvl = lvl.get(qn("w:ilvl"))
                if ilvl is None:
                    continue

                levels[int(ilvl)] = lvl

            abstract_levels[abstract_id] = levels

        # ------------------------------------------------------------------
        # 2. Build:
        #
        #     (numId, ilvl) -> NumberingInfo
        #
        # ------------------------------------------------------------------
        lookup: dict[tuple[int, int], NumberingInfo] = {}

        for num in numbering.findall(qn("w:num")):
            num_id_attr = num.get(qn("w:numId"))
            abstract_ref = num.find(qn("w:abstractNumId"))

            if num_id_attr is None or abstract_ref is None:
                continue

            num_id = int(num_id_attr)

            abstract_id_attr = abstract_ref.get(qn("w:val"))
            if abstract_id_attr is None:
                continue

            abstract_id = int(abstract_id_attr)

            levels = abstract_levels.get(abstract_id, {})
            if not levels:
                continue

            # --------------------------------------------------------------
            # A num can override properties of individual abstract levels.
            #
            # For now we support:
            #   - lvlOverride
            #   - startOverride
            #
            # A lvlOverride can also contain its own <w:lvl>, which should
            # replace the abstract level definition.
            # --------------------------------------------------------------
            overrides: dict[int, object] = {}

            for override in num.findall(qn("w:lvlOverride")):
                ilvl_attr = override.get(qn("w:ilvl"))
                if ilvl_attr is None:
                    continue

                ilvl = int(ilvl_attr)

                override_lvl = override.find(qn("w:lvl"))
                if override_lvl is not None:
                    overrides[ilvl] = override_lvl

            for ilvl, abstract_lvl in levels.items():
                lvl = overrides.get(ilvl, abstract_lvl)

                num_fmt = lvl.find(qn("w:numFmt")) # type: ignore oxml element
                lvl_text = lvl.find(qn("w:lvlText")) # type: ignore oxml element
                start = lvl.find(qn("w:start")) # type: ignore oxml element

                format_value = (
                    num_fmt.get(qn("w:val"))
                    if num_fmt is not None
                    else "none"
                )

                level_text = (
                    lvl_text.get(qn("w:val"))
                    if lvl_text is not None
                    else ""
                )

                start_value = (
                    int(start.get(qn("w:val")))
                    if start is not None and start.get(qn("w:val")) is not None
                    else 1
                )

                # Handle num-level startOverride.
                override = num.find(
                    f'{qn("w:lvlOverride")}[@{qn("w:ilvl")}="{ilvl}"]'
                )

                if override is not None:
                    start_override = override.find(qn("w:startOverride"))

                    if start_override is not None:
                        value = start_override.get(qn("w:val"))
                        if value is not None:
                            start_value = int(value)

                is_bullet = format_value == "bullet"

                is_alpha = format_value in {
                    "lowerLetter",
                    "upperLetter",
                }

                is_numeric = format_value in {
                    "decimal",
                    "decimalZero",
                }

                is_ordered = not is_bullet and format_value not in {
                    "none",
                    "bullet",
                }

                lookup[(num_id, ilvl)] = NumberingInfo(
                    num_id=num_id,
                    abstract_num_id=abstract_id,
                    level=ilvl,
                    format=format_value,
                    level_text=level_text,
                    start=start_value,
                    is_bullet=is_bullet,
                    is_ordered=is_ordered,
                    is_numeric=is_numeric,
                    is_alpha=is_alpha,
                )

        return lookup