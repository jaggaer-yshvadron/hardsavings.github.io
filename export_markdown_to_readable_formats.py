from __future__ import annotations

from dataclasses import dataclass
from html import escape
from pathlib import Path
import re


ROOT = Path("/home/yshvadro/Hard Savings")
SOURCES = [
    ROOT / "procure_to_pay_hard_savings_research_summary.md",
    ROOT / "value_articulation_hard_savings_inventory.md",
]


@dataclass
class Block:
    kind: str
    text: str


def parse_markdown(text: str) -> list[Block]:
    blocks: list[Block] = []
    lines = text.splitlines()

    for raw in lines:
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            blocks.append(Block("blank", ""))
            continue
        if stripped.startswith("### "):
            blocks.append(Block("h3", stripped[4:].strip()))
            continue
        if stripped.startswith("## "):
            blocks.append(Block("h2", stripped[3:].strip()))
            continue
        if stripped.startswith("# "):
            blocks.append(Block("h1", stripped[2:].strip()))
            continue
        if re.match(r"^\d+\.\s+", stripped):
            blocks.append(Block("ol", re.sub(r"^\d+\.\s+", "", stripped)))
            continue
        if stripped.startswith("- "):
            blocks.append(Block("ul", stripped[2:].strip()))
            continue
        blocks.append(Block("p", stripped))
    return blocks


def inline_html(text: str) -> str:
    return re.sub(r"`([^`]+)`", lambda m: f"<code>{escape(m.group(1))}</code>", escape(text))


def write_html(path: Path, blocks: list[Block]) -> None:
    parts = [
        "<!doctype html>",
        "<html lang=\"en\">",
        "<head>",
        "<meta charset=\"utf-8\">",
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
        f"<title>{escape(path.stem)}</title>",
        "<style>",
        "body{font-family:Georgia,'Times New Roman',serif;line-height:1.55;max-width:900px;margin:40px auto;padding:0 24px;color:#1f2937;background:#faf8f2;}",
        "h1,h2,h3{font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.2;}",
        "h1{font-size:2rem;margin-top:0;}",
        "h2{font-size:1.35rem;margin-top:1.8rem;border-top:1px solid #d1d5db;padding-top:1rem;}",
        "h3{font-size:1.05rem;margin-top:1.25rem;}",
        "p,li{font-size:1rem;}",
        "code{background:#ece7db;padding:.08rem .3rem;border-radius:4px;font-family:'Courier New',monospace;}",
        "ul,ol{margin:0.4rem 0 0.8rem 1.5rem;}",
        "</style>",
        "</head>",
        "<body>",
    ]

    list_mode: str | None = None
    for block in blocks:
        if block.kind not in {"ul", "ol"} and list_mode is not None:
            parts.append(f"</{list_mode}>")
            list_mode = None

        if block.kind == "blank":
            continue
        if block.kind == "h1":
            parts.append(f"<h1>{inline_html(block.text)}</h1>")
        elif block.kind == "h2":
            parts.append(f"<h2>{inline_html(block.text)}</h2>")
        elif block.kind == "h3":
            parts.append(f"<h3>{inline_html(block.text)}</h3>")
        elif block.kind == "p":
            parts.append(f"<p>{inline_html(block.text)}</p>")
        elif block.kind == "ul":
            if list_mode != "ul":
                if list_mode is not None:
                    parts.append(f"</{list_mode}>")
                parts.append("<ul>")
                list_mode = "ul"
            parts.append(f"<li>{inline_html(block.text)}</li>")
        elif block.kind == "ol":
            if list_mode != "ol":
                if list_mode is not None:
                    parts.append(f"</{list_mode}>")
                parts.append("<ol>")
                list_mode = "ol"
            parts.append(f"<li>{inline_html(block.text)}</li>")

    if list_mode is not None:
        parts.append(f"</{list_mode}>")

    parts.extend(["</body>", "</html>"])
    path.write_text("\n".join(parts), encoding="utf-8")


def rtf_escape(text: str) -> str:
    text = text.replace("\\", r"\\").replace("{", r"\{").replace("}", r"\}")
    return re.sub(r"`([^`]+)`", lambda m: r"\b " + m.group(1) + r"\b0 ", text)


def write_rtf(path: Path, blocks: list[Block]) -> None:
    parts = [
        r"{\rtf1\ansi\deff0",
        r"{\fonttbl{\f0 Times New Roman;}{\f1 Arial;}{\f2 Courier New;}}",
        r"\viewkind4\uc1\fs24",
    ]

    for block in blocks:
        text = rtf_escape(block.text)
        if block.kind == "blank":
            parts.append(r"\par")
        elif block.kind == "h1":
            parts.append(rf"\pard\sa200\b\f1\fs36 {text}\b0\fs24\f0\par")
        elif block.kind == "h2":
            parts.append(rf"\pard\sa160\b\f1\fs30 {text}\b0\fs24\f0\par")
        elif block.kind == "h3":
            parts.append(rf"\pard\sa120\b\f1\fs26 {text}\b0\fs24\f0\par")
        elif block.kind == "p":
            parts.append(rf"\pard\sa120 {text}\par")
        elif block.kind == "ul":
            parts.append(rf"\pard\li360\sa80 - {text}\par")
        elif block.kind == "ol":
            parts.append(rf"\pard\li360\sa80 {text}\par")

    parts.append("}")
    path.write_text("\n".join(parts), encoding="utf-8")


def main() -> None:
    for source in SOURCES:
        blocks = parse_markdown(source.read_text(encoding="utf-8"))
        write_html(source.with_suffix(".html"), blocks)
        write_rtf(source.with_suffix(".rtf"), blocks)
        print(f"Exported {source.name}")


if __name__ == "__main__":
    main()
