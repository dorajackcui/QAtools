"""Check repository Markdown links and the documented canonical CLI directory."""

from __future__ import annotations

from html.parser import HTMLParser
from pathlib import Path
import re
import subprocess
import sys
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]


def prose(text: str) -> str:
    return re.sub(r"^```[^\n]*\n.*?^```\s*$", "", text, flags=re.M | re.S)


def markdown_anchors(text: str) -> set[str]:
    anchors: set[str] = set()
    counts: dict[str, int] = {}
    for heading in re.findall(r"^#{1,6}\s+(.+?)\s*#*\s*$", prose(text), re.M):
        slug = re.sub(r"[^\w\- ]", "", heading.lower()).replace(" ", "-")
        count = counts.get(slug, 0)
        counts[slug] = count + 1
        anchors.add(f"{slug}-{count}" if count else slug)
    return anchors


class HtmlLinks(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []
        self.anchors: set[str] = set()

    def handle_starttag(self, tag, attrs):
        for name, value in attrs:
            if value and name in {"href", "src"}:
                self.links.append(value)
            if value and (name == "id" or tag == "a" and name == "name"):
                self.anchors.add(value)


def check() -> list[str]:
    listed = subprocess.check_output(
        ["git", "ls-files", "-z", "--cached", "--others", "--exclude-standard"], cwd=ROOT,
    ).decode("utf-8").split("\0")
    paths = sorted({ROOT / name for name in listed if name.endswith((".md", ".html"))})
    errors: list[str] = []
    for path in paths:
        if not path.is_file():
            continue  # A tracked file can be removed in the current working tree.
        source = path.read_text(encoding="utf-8")
        if path.suffix == ".md":
            links = re.findall(r"!?\[[^\]\n]*\]\((<[^>]+>|[^\s)]+)(?:\s+\"[^\"]*\")?\)", prose(source))
        else:
            html = HtmlLinks()
            html.feed(source)
            links = html.links
        for link in links:
            url = urlsplit(link.strip("<>"))
            if url.scheme or url.netloc:
                continue  # No network requests; remote sources aren't validated here.
            target = (path.parent / unquote(url.path)).resolve() if url.path else path
            if not target.exists():
                errors.append(f"{path.relative_to(ROOT)}: missing {link}")
                continue
            if url.fragment and target.suffix in {".md", ".html"}:
                text = target.read_text(encoding="utf-8")
                if target.suffix == ".md":
                    anchors = markdown_anchors(text)
                else:
                    html = HtmlLinks()
                    html.feed(text)
                    anchors = html.anchors
                if unquote(url.fragment) not in anchors:
                    errors.append(f"{path.relative_to(ROOT)}: missing anchor {link}")

    sys.path.insert(0, str(ROOT))
    from qatools.cli import COMMANDS

    guide = (ROOT / "docs/cli-usage.md").read_text(encoding="utf-8")
    directory = guide.split("## 命令目录\n", 1)[1].split("可用别名：", 1)[0]
    documented = re.findall(r"\| `qatools ([\w-]+)` \|", directory)
    expected = [command.name for command in COMMANDS]
    if documented != expected:
        errors.append(f"CLI command directory differs: documented={documented}, registered={expected}")
    return errors


if __name__ == "__main__":
    failures = check()
    for failure in failures:
        print(failure)
    print(f"Documentation check: {len(failures)} error(s).")
    raise SystemExit(bool(failures))
