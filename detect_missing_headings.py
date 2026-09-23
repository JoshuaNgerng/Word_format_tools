import re

from doc_manager import SpecialStyle

class HeadingDetection:
    HEADING_PATTERNS : list[tuple[re.Pattern, int | SpecialStyle]] = [
        (re.compile(r"^\d+\.\d+\.\d+\.\s+"), 3),   # 1.1.1
        (re.compile(r"^\d+\.0\s+"), 1),
        (re.compile(r"^\d+\.\d+\s+"), 2),         # 1.1
        (re.compile(r"^\d+\.\d+\.\b"), 2),          # 1.1
        (re.compile(r"^\d+\.\d+\.\s+"), 2),         # 1.1
        (re.compile(r"^\d+\.\d+\.\b"), 2),          # 1.1
        (re.compile(r"^\d+\.\b"), 1),           # 1.0
        (re.compile(r"^[a-zA-Z][.)]\s+"), SpecialStyle.Alphabet),
        (re.compile(r"^\s*[•*-]\s+"), SpecialStyle.Alphabet),
        (re.compile(r"^APPENDIX\s+"), SpecialStyle.Appendix)
    ]

    @classmethod
    def infer_heading_from_text(cls, text: str):
        text = text.strip()
        to_remove = ''
        val = None
        for pattern, v in cls.HEADING_PATTERNS:

            m = pattern.match(text)

            if not m:
                continue

            to_remove = m.group(0).strip()
            val = v
            break

        if isinstance(val, SpecialStyle):
            if val == SpecialStyle.Appendix:
                to_remove = ''

        # print(f'debug {val}, {to_remove} {text}')
        return val, to_remove
            