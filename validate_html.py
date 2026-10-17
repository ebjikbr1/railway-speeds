from html.parser import HTMLParser
from pathlib import Path
import re

root = Path(r"X:\My proekt\скоростя\html")
required = {"index.html"}
required |= {f"track-{track}-{kind}.html" for track in (1, 2) for kind in ("passenger", "freight", "empty", "18-100")}


class Checker(HTMLParser):
    def __init__(self):
        super().__init__()
        self.details = 0
        self.summaries = 0
        self.open_details = 0
        self.open_stations = 0
        self.closed_sections = 0
        self.links = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "details":
            self.details += 1
            classes = attrs.get("class", "").split()
            is_open = "open" in attrs
            if is_open:
                self.open_details += 1
            if "station" in classes and is_open:
                self.open_stations += 1
            if "section" in classes and not is_open:
                self.closed_sections += 1
        elif tag == "summary":
            self.summaries += 1
        elif tag == "a" and attrs.get("href"):
            self.links.append(attrs["href"])


missing = required - {p.name for p in root.glob("*.html")}
assert not missing, f"Missing pages: {sorted(missing)}"

for path in sorted(root.glob("*.html")):
    text = path.read_text(encoding="utf-8")
    assert "\ufffd" not in text, f"Replacement character in {path.name}"
    parser = Checker()
    parser.feed(text)
    if re.match(r"track-[12]-(passenger|freight|empty|18-100)\.html$", path.name):
        assert parser.details == parser.summaries > 0, f"Broken groups in {path.name}"
        assert parser.open_stations == 0, f"Stations must be collapsed in {path.name}"
        assert parser.closed_sections == 0, f"Non-station sections must be open in {path.name}"
        assert text.count('class="bar1 sticky"') == 1
        assert text.count('class="bar2 sticky"') == 1
        assert text.count('class="bar3 sticky"') == 1
        assert text.count('class="bar4 sticky"') == 1
        assert text.count('id="toggleNotes"') == 1
        assert '.note{width:30%;visibility:hidden;' in text
        assert '.notes-visible .note{visibility:visible}' in text
        assert 'data-segments="' in text
        assert "new URLSearchParams(location.search)" in text
    elif path.name == "index.html":
        assert text.count('data-group="section"') == 7
        assert text.count('data-group="track"') == 2
        assert text.count('data-group="train"') == 4
        assert 'id="go" disabled' in text
    for href in parser.links:
        if re.match(r"^[a-z0-9-]+\.html$", href):
            assert (root / href).exists(), f"Broken link {href} in {path.name}"
    print(f"OK {path.name}: details={parser.details}, links={len(parser.links)}")
