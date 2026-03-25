"""Patch chart OOXML after openpyxl.save(): fix c:delete=1 and stripped default-namespace <catAx>/<valAx>."""

import os
import re
import tempfile
import zipfile
from pathlib import Path

_RE_DEL = re.compile(r'<c:delete\s+val="1"\s*/>')
_RE_CAT = re.compile(r"<catAx>.*?</catAx>", re.DOTALL)
_RE_VAL = re.compile(r"<valAx>.*?</valAx>", re.DOTALL)


def _pair(block: str) -> tuple[str, str] | None:
    # Keep axis linkage intact when rebuilding blocks.
    a = re.search(r'<axId val="(\d+)"', block)
    c = re.search(r'<crossAx val="(\d+)"', block)
    return (a.group(1), c.group(1)) if a and c else None


def _axis(cat: bool, m: re.Match[str]) -> str:
    """Rebuild one axis XML block: OpenPyXL leaves <catAx>/<valAx> too small; Excel needs tickLblPos, numFmt, etc."""
    b = m.group(0)
    p = _pair(b)
    if not p:
        return b
    i, x = p
    if cat:
        # Bottom category axis: restore tick labels + crosses; keep axId/crossAx/lblOffset from the stub.
        lo = re.search(r'<lblOffset val="(\d+)"', b)
        off = lo.group(1) if lo else "100"
        return (
            f'<catAx><axId val="{i}"/><scaling><orientation val="minMax"/></scaling>'
            f'<delete val="0"/><axPos val="b"/><numFmt formatCode="General" sourceLinked="0"/>'
            f'<majorTickMark val="none"/><minorTickMark val="none"/><tickLblPos val="nextTo"/>'
            f'<crossAx val="{x}"/><crosses val="autoZero"/><auto val="1"/><lblAlgn val="ctr"/>'
            f'<lblOffset val="{off}"/><noMultiLvlLbl val="1"/></catAx>'
        )
    # Left value axis: keep % format if present; keep gridlines flag if it was there.
    nf = re.search(r'<numFmt formatCode="([^"]*)"', b)
    sl = re.search(r'sourceLinked="(\d+)"', b)
    fmt, slv = (nf.group(1) if nf else "0%"), (sl.group(1) if sl else "1")
    g = "<majorGridlines/>" if "majorGridlines" in b else ""
    return (
        f'<valAx><axId val="{i}"/><scaling><orientation val="minMax"/></scaling>'
        f'<delete val="0"/><axPos val="l"/>{g}'
        f'<numFmt formatCode="{fmt}" sourceLinked="{slv}"/>'
        f'<majorTickMark val="none"/><minorTickMark val="none"/><tickLblPos val="nextTo"/>'
        f'<crossAx val="{x}"/><crosses val="autoZero"/><crossBetween val="midCat"/></valAx>'
    )


def _patch_xml(s: str) -> tuple[str, bool]:
    # Prefixed namespace case: only delete-flag is wrong.
    t = _RE_DEL.sub('<c:delete val="0"/>', s)
    ch = t != s
    # Default namespace case: OpenPyXL can strip full axis details.
    if "<c:catAx" not in t and "<catAx>" in t:
        u = _RE_VAL.sub(lambda m: _axis(False, m), _RE_CAT.sub(lambda m: _axis(True, m), t))
        if u != t:
            ch = True
            t = u
    return t, ch


def repair_chart_axes_in_xlsx_path(filepath: str | Path) -> bool:
    path = Path(filepath)
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    mod = False
    fd, name = tempfile.mkstemp(suffix=path.suffix, dir=path.parent)
    os.close(fd)
    tmp = Path(name)
    try:
        # xlsx is a zip; patch only chart XML entries and copy everything else as-is.
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.startswith("xl/charts/chart") and info.filename.endswith(".xml"):
                    text, c = _patch_xml(data.decode("utf-8"))
                    if c:
                        mod = True
                        data = text.encode("utf-8")
                zout.writestr(info, data)
        if mod:
            os.replace(tmp, path)
        else:
            tmp.unlink(missing_ok=True)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    return mod
