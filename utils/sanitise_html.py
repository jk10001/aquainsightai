from __future__ import annotations

import re
from typing import Optional

from bs4 import BeautifulSoup
import bleach

# ────────────────────────────────────────────────────────────────
# Config: allowed tags & attributes (rich text kept, images allowed)
# ────────────────────────────────────────────────────────────────

ALLOWED_TAGS = [
    # text structure & formatting
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "p",
    "div",
    "span",
    "br",
    "hr",
    "ul",
    "ol",
    "li",
    "strong",
    "em",
    "b",
    "i",
    "u",
    "s",
    "sub",
    "sup",
    "blockquote",
    "code",
    "pre",
    "table",
    "thead",
    "tbody",
    "tr",
    "th",
    "td",
    "img",
    # links (kept, but href restricted to local-only below)
    "a",
]

# Attribute policy: provide a callable so we can enforce local-only src/href and scrub styles.
# For most tags we allow "style" (layout is fine), but we remove *dangerous* CSS constructs below.
ALLOWED_GLOBAL_ATTRS = {"title", "aria-label", "role"}
ALLOWED_STYLE_ATTR = True  # keep inline style, but scrub dangerous constructs
ALLOWED_LINK_ATTRS = {"href", "title"}
ALLOWED_IMG_ATTRS = {"src", "alt", "title", "style"}

# Dangerous CSS patterns we will strip entirely if present
DANGEROUS_CSS_PATTERNS = re.compile(
    r"""url\(|@import|expression\(|behavior\(|-moz-binding|javascript:""",
    re.IGNORECASE,
)

# Local filename/path rule:
# - relative only (no scheme, no //, no leading /)
# - no traversal, no backslashes
# - simple subpaths allowed (e.g. "imgs/chart.png")
# - common image extensions
LOCAL_IMG_RE = re.compile(r"^[\w.\-\/]+\.(png|jpe?g|gif|webp|svg)$", re.IGNORECASE)

# Relative-only URI for anchors (no external links)
LOCAL_URI_RE = re.compile(r"^(?!([a-z][a-z0-9+.\-]*:)?\/\/)", re.IGNORECASE)


def _is_local_image_path(value: str) -> bool:
    if not value:
        return False
    if value.startswith("//"):  # protocol-relative
        return False
    if re.match(
        r"^[a-z][a-z0-9+.\-]*:", value, re.IGNORECASE
    ):  # has a scheme (http:, data:, blob:, etc.)
        return False
    if value.startswith("/"):  # absolute path
        return False
    if ".." in value or "\\" in value:
        return False
    return bool(LOCAL_IMG_RE.match(value))


def _is_local_uri(value: str) -> bool:
    # Accept empty or fragment-only (“#id”) or relative paths (“report.html”, “section/part”)
    if not value:
        return False
    if value.startswith("#"):
        return True
    if value.startswith("//"):
        return False
    if re.match(r"^[a-z][a-z0-9+.\-]*:", value, re.IGNORECASE):  # has a scheme
        return False
    if value.startswith(
        "/"
    ):  # reject absolute root paths for safety (optional; relax if you serve from root)
        return False
    if ".." in value or "\\" in value:
        return False
    return True


def _scrub_style(value: str) -> Optional[str]:
    """Return the style string if it looks safe; otherwise None to drop."""
    if value is None:
        return None
    # Security-only scrub: drop the entire style if any dangerous construct appears.
    if DANGEROUS_CSS_PATTERNS.search(value):
        return None
    return value


def _attribute_filter(tag: str, name: str, value: str) -> Optional[str]:
    """
    bleach Cleaner(attributes=callable) signature:
    - Return a (possibly modified) string to keep the attribute,
    - or None to drop it.
    """
    t = tag.lower()
    n = name.lower()

    # Allow some universal safe attrs
    if n in ALLOWED_GLOBAL_ATTRS:
        return value

    # Styles: keep layout but remove dangerous CSS
    if n == "style" and ALLOWED_STYLE_ATTR:
        return _scrub_style(value)

    # IMG: enforce local-only src; scrub style separately above
    if t == "img":
        if n in {"alt", "title"}:
            return value
        if n == "src":
            return value if _is_local_image_path(value) else None
        if n == "style":
            return _scrub_style(value)
        return None  # drop any other img attrs (onerror, onclick, etc.)

    # A: allow only local hrefs (no external beacons)
    if t == "a":
        if n == "href":
            return value if _is_local_uri(value) else None
        if n in {"title"}:
            return value
        if n == "style" and ALLOWED_STYLE_ATTR:
            return _scrub_style(value)
        return None

    # Other tags: permit style only (already scrubbed), or drop unknown attrs
    if n == "style" and ALLOWED_STYLE_ATTR:
        return _scrub_style(value)

    # Everything else: drop
    return None


# Tags to REMOVE ENTIRELY with their contents (no inner text kept)
NUKE_TAGS = {
    "script",
    "style",
    "iframe",
    "object",
    "embed",
    "link",
    "meta",
    "base",
    "form",
    "input",
    "button",
    "textarea",
    "select",
}


def _nuke_dangerous_nodes(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all(NUKE_TAGS):
        tag.decompose()
    return str(soup)


def sanitise_html(unsafe_html: str) -> str:
    """
    - Removes dangerous nodes entirely (script/style/iframe/etc.)
    - Keeps rich text tags (h1/p/div/ul/li/strong/etc.)
    - Allows <img> but only with local-only src (e.g., 'chart.png' or 'imgs/chart.png')
    - Keeps inline style but strips security-risky constructs (url(), @import, expression(), behavior(), -moz-binding, javascript:)
    - Keeps <a> but only local hrefs (no http/https/data/blob/etc.)
    """
    # 1) Remove whole dangerous elements and their contents
    precleaned = _nuke_dangerous_nodes(unsafe_html)

    # 2) Bleach sanitize with our attribute filter
    cleaner = bleach.Cleaner(
        tags=ALLOWED_TAGS,
        attributes=_attribute_filter,
        strip=True,  # strip disallowed tags (their *content* already nuked for dangerous ones)
        strip_comments=True,
    )

    cleaned = cleaner.clean(precleaned)
    return cleaned
