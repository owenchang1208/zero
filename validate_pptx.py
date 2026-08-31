"""Validate generated PPTX structure and editability assumptions."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def validate_pptx(path: Path) -> None:
    """Validate the generated pptx file.

    Args:
        path: Path to the pptx file.

    Returns:
        None.

    Raises:
        FileNotFoundError: If the pptx file does not exist.
        ValueError: If the pptx file is invalid or violates checks.
    """
    if not path.exists():
        raise FileNotFoundError(f"PPTX not found: {path}")

    with zipfile.ZipFile(path, "r") as zf:
        names = set(zf.namelist())
        if "[Content_Types].xml" not in names:
            raise ValueError("Invalid PPTX: missing [Content_Types].xml")
        if "ppt/slides/slide1.xml" not in names:
            raise ValueError("Invalid PPTX: missing ppt/slides/slide1.xml")

        slide_xml = zf.read("ppt/slides/slide1.xml")
        root = ET.fromstring(slide_xml)

        shapes = root.findall(f".//{P_NS}sp")
        text_runs = root.findall(f".//{A_NS}t")
        pictures = root.findall(f".//{P_NS}pic")

        if len(shapes) < 40:
            raise ValueError(f"Too few editable shapes: {len(shapes)}")
        if len(text_runs) < 40:
            raise ValueError(f"Too few text runs: {len(text_runs)}")
        if len(pictures) > 6:
            raise ValueError(
                "Suspiciously high picture count; possible raster-image embedding"
            )

    print("Validation passed")
    print(f"- Editable shapes: {len(shapes)}")
    print(f"- Text runs: {len(text_runs)}")
    print(f"- Pictures: {len(pictures)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--file", default="output/aegisops_mvp_redraw.pptx", help="PPTX file path"
    )
    args = parser.parse_args()
    validate_pptx(Path(args.file))
