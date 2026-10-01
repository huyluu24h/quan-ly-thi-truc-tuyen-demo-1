"""Convert DOCX equations to normalized LaTeX without changing the exam parser.

MathType OLE equations are decoded with MTEF-py. Native Word equations (OMML)
are converted by a small, defensive converter. LaTeX is the only stored formula
representation; local MathJax renders it consistently for every web screen.
Raster/vector previews are not retained as formula display data.
"""

from __future__ import annotations

import contextlib
import io
import re
import sys
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree


M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
O_NS = "urn:schemas-microsoft-com:office:office"
NS = {"m": M_NS, "r": R_NS, "o": O_NS}


def _escape_text(value: str) -> str:
    value = value or ""
    value = value.replace("\\", r"\backslash ")
    for char, replacement in {
        "{": r"\{", "}": r"\}", "#": r"\#", "%": r"\%",
        "&": r"\&", "_": r"\_", "−": "-", "×": r"\times ",
        "÷": r"\div ", "≤": r"\le ", "≥": r"\ge ", "≠": r"\ne ",
        "∞": r"\infty ", "π": r"\pi ", "√": r"\sqrt{}",
    }.items():
        value = value.replace(char, replacement)
    return value


def _children(node, local_name: str):
    return node.xpath("./m:" + local_name, namespaces=NS)


def _first(node, local_name: str):
    values = _children(node, local_name)
    return values[0] if values else None


def _omml(node) -> str:
    name = etree.QName(node).localname
    if name in {"oMath", "oMathPara", "e", "num", "den", "sub", "sup", "deg"}:
        return "".join(_omml(child) for child in node if etree.QName(child).localname not in {"ctrlPr", "oMathParaPr"})
    if name == "r":
        return "".join(_escape_text(text) for text in node.xpath("./m:t/text()", namespaces=NS))
    if name == "t":
        return _escape_text(node.text or "")
    if name == "f":
        # Display-style fractions match the original MathType proportions and
        # remain readable beside 13.5pt question text.
        return r"\dfrac{" + _omml(_first(node, "num")) + "}{" + _omml(_first(node, "den")) + "}"
    if name in {"sSup", "sSub", "sSubSup"}:
        base = _omml(_first(node, "e"))
        sub = _first(node, "sub"); sup = _first(node, "sup")
        return "{" + base + "}" + (("_{" + _omml(sub) + "}") if sub is not None else "") + (("^{" + _omml(sup) + "}") if sup is not None else "")
    if name == "rad":
        base = _omml(_first(node, "e")); degree = _first(node, "deg")
        degree_latex = _omml(degree) if degree is not None else ""
        return (r"\sqrt[" + degree_latex + "]{" + base + "}") if degree_latex else (r"\sqrt{" + base + "}")
    if name == "d":
        base = _omml(_first(node, "e"))
        beg = "("; end = ")"
        vals = node.xpath("./m:dPr/m:begChr/@m:val", namespaces=NS)
        if vals: beg = vals[0]
        vals = node.xpath("./m:dPr/m:endChr/@m:val", namespaces=NS)
        if vals: end = vals[0]
        return r"\left" + _escape_text(beg) + " " + base + r"\right" + _escape_text(end)
    if name == "nary":
        op = "∑"
        vals = node.xpath("./m:naryPr/m:chr/@m:val", namespaces=NS)
        if vals: op = vals[0]
        command = {"∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∬": r"\iint", "∭": r"\iiint"}.get(op, _escape_text(op))
        sub = _first(node, "sub"); sup = _first(node, "sup"); base = _first(node, "e")
        return command + (("_{" + _omml(sub) + "}") if sub is not None else "") + (("^{" + _omml(sup) + "}") if sup is not None else "") + " " + (_omml(base) if base is not None else "")
    if name in {"bar", "acc"}:
        base = _omml(_first(node, "e"))
        char = "¯"
        vals = node.xpath("./m:accPr/m:chr/@m:val | ./m:barPr/m:pos/@m:val", namespaces=NS)
        if vals: char = vals[0]
        command = r"\vec" if char in {"⃗", "→"} else (r"\underline" if char == "bot" else r"\overline")
        return command + "{" + base + "}"
    if name == "func":
        fname = _omml(_first(node, "fName")); base = _omml(_first(node, "e"))
        return fname + " " + base
    if name == "fName":
        return "".join(_omml(child) for child in node)
    if name == "m":
        rows = []
        for row in _children(node, "mr"):
            rows.append(" & ".join(_omml(cell) for cell in _children(row, "e")))
        return r"\begin{matrix}" + r" \\ ".join(rows) + r"\end{matrix}"
    if name == "eqArr":
        return r"\begin{aligned}" + r" \\ ".join(_omml(e) for e in _children(node, "e")) + r"\end{aligned}"
    return "".join(_omml(child) for child in node if etree.QName(child).localname not in {"rPr", "ctrlPr"} )


def omml_to_latex(node) -> str | None:
    try:
        latex = re.sub(r"\s+", " ", _omml(node)).strip()
        return latex or None
    except Exception:
        return None


def _normalise_latex(value: str) -> str | None:
    """Return delimiter-free, browser-safe LaTeX.

    MTEF-py intentionally returns ``$ ... $`` and leaves a few Unicode maths
    characters untouched.  The application stores only the body because its
    templates supply MathJax delimiters themselves.
    """
    latex = (value or "").strip()
    # Older versions and MTEF-py can return their own delimiters.  The database
    # stores only the equation body; templates add \(...\) or \[...\].
    delimiter_pairs = (("$$", "$$"), (r"\[", r"\]"), (r"\(", r"\)"), ("$", "$"))
    changed = True
    while changed:
        changed = False
        for opening, closing in delimiter_pairs:
            if (len(latex) >= len(opening) + len(closing)
                    and latex.startswith(opening) and latex.endswith(closing)):
                latex = latex[len(opening):-len(closing)].strip()
                changed = True
                break
    latex = re.sub(r"[\uE000-\uF8FF]", "", latex)
    latex = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", latex)
    for char, command in {
        "−": "-",
        "α": r"\alpha ",
        "Δ": r"\Delta ",
        "∞": r"\infty ",
        "π": r"\pi ",
        "×": r"\times ",
        "÷": r"\div ",
        "∈": r"\in ",
        "∉": r"\notin ",
        "∪": r"\cup ",
        "∩": r"\cap ",
        "→": r"\to ",
        "⇒": r"\Rightarrow ",
        "⇔": r"\Leftrightarrow ",
        "°": r"^{\circ}",
        "⊥": r"\perp ",
        "≤": r"\le ",
        "≥": r"\ge ",
        "≠": r"\ne ",
    }.items():
        latex = latex.replace(char, command)

    # Canonical LaTeX used by the application.  MTEF-py emits forms such as
    # ``\frac { a } { b }``; MathJax accepts them, but normalising to
    # ``\dfrac{a}{b}`` makes fractions consistently display-style and avoids
    # browser/parser differences.  Radicals always keep an explicit group.
    latex = re.sub(r"\\(?:dfrac|tfrac|frac)\s*(?=\{)", r"\\dfrac", latex)
    # MTEF-py writes ordinary square roots as ``\sqrt[] {x}``.  An empty
    # optional degree is not canonical TeX and can produce merror in MathJax,
    # so collapse it to ``\sqrt{x}``; keep a real n-th-root degree intact.
    latex = re.sub(r"\\sqrt\s*\[\s*\]\s*\{", r"\\sqrt{", latex)
    latex = re.sub(
        r"\\sqrt\s*\[\s*([^\]]+?)\s*\]\s*\{",
        lambda match: r"\sqrt[" + match.group(1).strip() + "]{",
        latex,
    )
    latex = re.sub(r"\\sqrt\s*\{", r"\\sqrt{", latex)
    # Véc-tơ xác định bởi hai điểm cần mũi tên phủ hết tên. ``\vec`` vẫn phù
    # hợp với một ký hiệu đơn như u, v; còn AB, MN, SA... dùng dạng dài.
    vector_letters = r"[A-Za-zΑ-ω]{2,}"
    latex = re.sub(
        rf"\\vec\s*\{{\s*({vector_letters})\s*\}}",
        lambda match: r"\overrightarrow{" + match.group(1) + "}",
        latex,
    )
    # MTEF-py có thể trả cùng dấu véc-tơ bằng hai biến thể \overset.
    latex = re.sub(
        rf"\\overset\s*\\rightarrow\s*\{{\s*({vector_letters})\s*\}}",
        lambda match: r"\overrightarrow{" + match.group(1) + "}",
        latex,
    )
    latex = re.sub(
        rf"\\overset\s*\{{\s*\\rightarrow\s*\}}\s*\{{\s*({vector_letters})\s*\}}",
        lambda match: r"\overrightarrow{" + match.group(1) + "}",
        latex,
    )
    # Common function names from MTEF are plain italic letters. Canonical
    # operator commands give the same upright typography as MathType.
    # PDF/MathType exporters also wrap lim/min/max in ``\mathrm`` or
    # ``\operatorname``; collapse those variants before processing plain text.
    latex = re.sub(
        r"\\(?:operatorname|mathrm|rm)\s*\{\s*(lim|min|max)\s*\}",
        lambda match: "\\" + match.group(1),
        latex,
    )
    latex = re.sub(
        r"(?<![\\A-Za-z])limm(?=\s*(?:_|\{))",
        r"\\lim",
        latex,
    )
    latex = re.sub(
        r"(?<![\\A-Za-z{])(arcsin|arccos|arctan|sin|cos|tan|cot|log|ln|lim|min|max)(?=\b|_)",
        lambda match: "\\" + match.group(1),
        latex,
    )
    # MTEF's legacy ``\rm{...}`` form is accepted by some TeX engines but is
    # not a stable web-math representation. Convert it to explicit
    # ``\mathrm{...}``, matching the upright unit text used by MathType.
    latex = re.sub(
        r"\{\s*\\rm\s*\{\s*([^{}]*?)\s*\}\s*\}",
        lambda match: r"\mathrm{" + match.group(1).strip() + "}",
        latex,
    )
    latex = re.sub(
        r"\\rm\s*\{\s*([^{}]*?)\s*\}",
        lambda match: r"\mathrm{" + match.group(1).strip() + "}",
        latex,
    )
    # MathType places the condition of lim/min/max below the operator, even
    # when the formula is embedded in a line of text. ``\limits`` preserves
    # that layout in MathJax instead of moving the condition to the right.
    latex = re.sub(
        r"\\(lim|min|max)(?![A-Za-z])\s*(?:\\limits\s*)?_",
        lambda match: "\\" + match.group(1) + r"\limits_",
        latex,
    )
    # Vietnamese school documents commonly abbreviate cực đại/cực tiểu as
    # y_CĐ and y_CT. Keep the abbreviation upright and entirely in the
    # subscript, matching the source MathType object.
    latex = re.sub(
        r"_\s*\{\s*(CĐ|CT|CD)\s*\}",
        lambda match: r"_{\mathrm{" + match.group(1) + "}}",
        latex,
    )
    latex = re.sub(
        r"_\s*(CĐ|CT|CD)(?![A-Za-zÀ-ỹ])",
        lambda match: r"_{\mathrm{" + match.group(1) + "}}",
        latex,
    )
    latex = re.sub(r"\\left\s+([()\[\]|.])", r"\\left\1", latex)
    latex = re.sub(r"\\right\s+([()\[\]|.])", r"\\right\1", latex)
    latex = re.sub(r"\\left\s+(\\[{}])", r"\\left\1", latex)
    latex = re.sub(r"\\right\s+(\\[{}])", r"\\right\1", latex)
    latex = re.sub(r"([_^])\s*\{", r"\1{", latex)
    latex = re.sub(r"[ \t\r\n]+", " ", latex).strip()
    if not latex or "¢" in latex or "�" in latex:
        return None
    # Literal delimiters \{ and \} are not grouping braces.
    open_groups = len(re.findall(r"(?<!\\)\{", latex))
    close_groups = len(re.findall(r"(?<!\\)\}", latex))
    if open_groups != close_groups:
        return None
    # Count delimiter commands only. A raw substring count treats the
    # ``\right`` inside ``\rightarrow`` as a closing delimiter and rejects
    # valid MathType vector formulas such as ``\overset\rightarrow{AB}``.
    left_count = len(re.findall(r"\\left(?![A-Za-z])", latex))
    right_count = len(re.findall(r"\\right(?![A-Za-z])", latex))
    if left_count != right_count:
        return None
    return latex


def normalise_latex(value: str) -> str | None:
    """Public normaliser shared by uploads, old database rows and manual input."""
    return _normalise_latex(value)


def _legacy_prime_char(node) -> bool:
    """MathType/Euclid Symbol stores a prime as MTCode U+00A2.

    U+00A2 is *not* a cent sign in this context.  It is the glyph position for
    the prime character in the legacy Symbol font used by the source document.
    """
    try:
        from mtef_py.record import RecordType
        return node.tag == RecordType.CHAR and node.value.mtcode == 0x00A2
    except Exception:
        return False


def _translate_legacy_derivative(equation) -> str | None:
    """Repair derivative expressions that upstream MTEF-py nests incorrectly.

    In some Equation Editor/MathType OLE objects the prime is encoded as the
    first item of a superscript slot using the legacy Symbol font.  Upstream
    treats it as a literal cent sign and keeps the remainder of the equation in
    that superscript.  A quotient can consequently lose its denominator and
    ``makeLatex`` raises ``IndexError`` (the user's Question 10B).

    This walker closes such a prime slot after its first glyph and propagates a
    displaced LINE back to the nearest fraction as its denominator.
    """
    from mtef_py.record import RecordType, SelectorType

    def translate_node(node):
        # Returns (latex, displaced slot lines).
        if node.tag in (RecordType.ROOT, RecordType.LINE):
            return translate_many(node.children)
        if node.tag == RecordType.CHAR:
            if _legacy_prime_char(node):
                return r"^{\prime}", []
            text, _ = equation.makeLatex(node)
            # In these derivative objects a full-stop glyph from the function
            # font is MathType's centred multiplication dot.
            if node.value.mtcode == 0x002E and node.value.typeface == 130:
                text = r"\cdot "
            return text, []
        if node.tag == RecordType.TMPL:
            children = list(node.children or [])
            if node.value.selector == SelectorType.tmSUP and len(children) >= 2:
                script = children[1]
                if (script.tag == RecordType.LINE and script.children and
                        _legacy_prime_char(script.children[0])):
                    tail, displaced = translate_many(script.children[1:])
                    displaced.extend(
                        child for child in children[2:]
                        if child.tag == RecordType.LINE
                    )
                    return r"^{\prime}" + tail, displaced
            if node.value.selector == SelectorType.tmFRACT and children:
                numerator, displaced = translate_node(children[0])
                if len(children) >= 2:
                    denominator_node = children[1]
                elif displaced:
                    denominator_node = displaced.pop(0)
                else:
                    raise ValueError("MTEF fraction has no denominator slot")
                denominator, more = translate_node(denominator_node)
                displaced.extend(more)
                return (r"\dfrac{" + numerator + "}{" + denominator + "}",
                        displaced)
        text, _ = equation.makeLatex(node)
        return text, []

    def translate_many(nodes):
        parts = []
        displaced = []
        for node in nodes:
            text, extra = translate_node(node)
            parts.append(text)
            displaced.extend(extra)
        return "".join(parts), displaced

    try:
        latex, displaced = translate_node(equation.ast)
        if displaced:
            return None
        return latex
    except Exception:
        return None


def _mtef_to_latex(data: bytes, vendor_dir: Path, return_details=False):
    try:
        parent = str(vendor_dir.parent)
        if parent not in sys.path:
            sys.path.insert(0, parent)
        from mtef_py.mtef import MTEF
        # The upstream decoder prints diagnostic lines; a web upload must stay quiet.
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            equation, error = MTEF.OpenBytes(data)
            if error or equation is None:
                return None
            # First use the upstream translation.  If it exposes the legacy
            # prime marker, rebuild just that MTEF tree with the repair above.
            try:
                latex = equation.Translate() or ""
            except IndexError:
                latex = ""
            repaired_legacy = not latex or "¢" in latex
            if repaired_legacy:
                latex = _translate_legacy_derivative(equation) or ""
        normalized = _normalise_latex(latex)
        if return_details:
            return normalized, repaired_legacy
        return normalized
    except Exception:
        return (None, False) if return_details else None


def _relationship_targets(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
    except Exception:
        return {}
    result = {}
    for rel in root.findall("{%s}Relationship" % PKG_REL_NS):
        target = rel.get("Target", "")
        if rel.get("Id") and target:
            result[rel.get("Id")] = str(PurePosixPath("word") / target)
    return result


def extract_docx_equations(path, vendor_dir=None) -> list[dict]:
    """Return equations in document order as ``{'latex': ..., 'display': ...}``."""
    vendor_dir = Path(vendor_dir or Path(__file__).parent / "vendor" / "mtef_py")
    equations = []
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        relationships = _relationship_targets(archive)
        # Select top-level OMML nodes to avoid emitting nested elements twice.
        nodes = root.xpath(".//m:oMath[not(ancestor::m:oMath)] | .//o:OLEObject", namespaces=NS)
        for node in nodes:
            local = etree.QName(node).localname
            latex = None
            display = "inline"
            if local == "oMath":
                latex = omml_to_latex(node)
                display = "block" if node.getparent() is not None and etree.QName(node.getparent()).localname == "oMathPara" else "inline"
            else:
                rel_id = node.get("{%s}id" % R_NS)
                target = relationships.get(rel_id, "")
                legacy_repaired = False
                if target in archive.namelist():
                    latex, legacy_repaired = _mtef_to_latex(
                        archive.read(target), vendor_dir, return_details=True
                    )
                # The VML preview lives beside the OLE object in w:object.
                obj = node
                while obj is not None and etree.QName(obj).localname != "object":
                    obj = obj.getparent()
                previews = obj.xpath(".//v:imagedata/@r:id", namespaces={**NS, "v": "urn:schemas-microsoft-com:vml"}) if obj is not None else []
                preview_target = relationships.get(previews[0], "") if previews else ""
            equations.append({
                "latex": latex,
                "display": display,
                "preview_target": preview_target if local != "oMath" else "",
                "legacy_repaired": legacy_repaired if local != "oMath" else False,
            })
    return equations


def replace_formula_images_with_latex(block_lines, equations):
    """Replace every Word formula image with canonical, delimiter-free LaTeX.

    Only real figures remain images.  Formula previews are deliberately not
    copied to the result, so teacher/student screens have one uniform MathJax
    rendering path and cannot mix differently sized raster/vector formulas.
    """
    index = 0
    replaced = 0
    for line in block_lines:
        for pos, block in enumerate(line):
            if block.get("type") == "image" and block.get("role") == "formula":
                if index < len(equations):
                    equation = equations[index]
                    if equation.get("latex"):
                        line[pos] = {
                            "type": "latex",
                            "value": equation["latex"],
                            "display": equation.get("display", "inline"),
                        }
                        replaced += 1
                index += 1
    if index != len(equations) or replaced != len(equations):
        raise ValueError(
            "Số vị trí công thức Word không khớp dữ liệu MTEF/OMML "
            f"({index} vị trí, {len(equations)} công thức, {replaced} đã đổi)."
        )
    return block_lines
