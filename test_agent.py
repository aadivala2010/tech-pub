#!/usr/bin/env python3
"""Self-check for the document revision agent.  Run:  python test_agent.py

No AI calls — exercises the file-handling logic only, which is where the
crashes were. Needs current_database/DOC1_RevA.docx and images_database/.
"""

import os
import shutil
import tempfile
import zipfile
from pathlib import Path

from agent import (
    best_gallery_match,
    increment_revision,
    index_gallery,
    replace_docx_images,
    safe_output_path,
    _extract_text_from_docx,
)

ROOT = Path(__file__).parent
SRC = ROOT / "current_database" / "DOC1_RevA.docx"


def _copy(tmp):
    dst = Path(tmp) / "doc.docx"
    shutil.copy2(SRC, dst)
    return str(dst)


def _media(path):
    with zipfile.ZipFile(path) as z:
        return {e: len(z.read(e)) for e in z.namelist() if e.startswith("word/media/")}


def test_gallery_match():
    g = index_gallery(str(ROOT / "images_database"))
    assert g, "gallery should not be empty"
    assert best_gallery_match(g, "Tecnord"), "exact stem must match"
    assert best_gallery_match(g, "Tecnord regulated valve"), "substring must match"
    assert best_gallery_match(g, "no-such-part-9999") is None
    assert best_gallery_match({}, "Tecnord") is None


def test_unmatched_image_does_not_crash():
    """The bug: no gallery match left new_img_path None, which reached open()
    and raised TypeError — and run_pipeline then deleted the saved text work."""
    g = index_gallery(str(ROOT / "images_database"))
    with tempfile.TemporaryDirectory() as tmp:
        doc = _copy(tmp)
        before = _media(doc)
        changes = [{"old_part": "ABC", "new_part": "DEF", "new_image": "DEF",
                    "old_specs": {}, "new_specs": {}}]
        assert replace_docx_images(doc, changes, g, lambda m: None) is False
        assert _media(doc) == before, "images must be untouched when nothing matched"
        assert _extract_text_from_docx(doc), "document must still be readable"


def test_matched_image_is_swapped():
    g = index_gallery(str(ROOT / "images_database"))
    tecnord = ROOT / "images_database" / "Tecnord.jpg"
    with tempfile.TemporaryDirectory() as tmp:
        doc = _copy(tmp)
        changes = [{"old_part": "Husco", "new_part": "Tecnord", "new_image": "Tecnord",
                    "old_specs": {}, "new_specs": {}}]
        assert replace_docx_images(doc, changes, g, lambda m: None) is True
        sizes = _media(doc)
        assert set(sizes.values()) == {tecnord.stat().st_size}, sizes


def test_revision_and_collision():
    assert increment_revision("DOC1_RevA.docx") == "DOC1_RevB.docx"
    assert increment_revision("DOC1_RevZ.docx") == "DOC1_RevAA.docx"
    assert increment_revision("plain.docx") == "plain_revA.docx"
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / "DOC1_RevB.docx"
        assert safe_output_path(target) == target
        target.write_bytes(b"x")
        assert safe_output_path(target).name == "DOC1_RevB_2.docx"


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
            print(f"  ok  {name}")
    print("\nAll checks passed.")
