"""Generate editable AegisOps MVP architecture slide."""

from pathlib import Path

from pptx_redraw.pipeline import generate_pptx


if __name__ == "__main__":
    output = Path("output/aegisops_mvp_redraw.pptx")
    generate_pptx(output)
    print(f"Generated: {output.resolve()}")
