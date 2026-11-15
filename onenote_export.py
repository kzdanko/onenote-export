#!/usr/bin/env python3
"""onenote-export: copy OneNote notebooks to a local Markdown tree.

Reads a personal Microsoft account's OneNote notebooks through Microsoft
Graph and writes one Markdown file per page, in folders that mirror the
notebook / section group / section / page hierarchy.

Commands:
    login   sign in with a device code and cache the token
    list    print the notebook hierarchy and the paths that would be written
    sync    export new and changed pages (incremental, resumable)
"""
from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import posixpath
import random
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlparse

import requests
from bs4 import BeautifulSoup
from markdownify import markdownify

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_HOST = "graph.microsoft.com"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.Read"]
CONFIG_DIR = Path(
    os.environ.get("ONENOTE_EXPORT_CONFIG_DIR")
    or Path.home() / ".config" / "onenote-export"
)
STATE_DIR = ".onenote-export"
ASSETS_DIR = "_assets"


# --------------------------------------------------------------------------
# Names and paths
# --------------------------------------------------------------------------

_WINDOWS_RESERVED = {
    "con", "prn", "aux", "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}


def sanitize_name(name: str, fallback: str = "Untitled", max_len: int = 120) -> str:
    """Make a string safe as a file or folder name on Windows, Linux and macOS.

    Accented letters are kept. Only characters that Windows forbids, or that
    break Obsidian links (# ^ [ ]), are replaced or removed.
    """
    name = unicodedata.normalize("NFC", name or "")
    name = re.sub(r"[\x00-\x1f\x7f]", " ", name)
    name = re.sub(r"[\\/:|]", "-", name)
    name = re.sub(r'[*?"<>#^\[\]]', "", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    if len(name) > max_len:
        name = name[:max_len].rstrip(" .")
    if not name:
        name = fallback
    if name.split(".")[0].lower() in _WINDOWS_RESERVED:
        name = "_" + name
    return name


def unique_names(entries, reserved=()):
    """Give each entry a name that is unique ignoring case.

    entries: iterable of (ident, desired_name, sort_key). The entry with the
    smallest sort key keeps the plain name; later ones get " (2)", " (3)".
    Sorting by creation time keeps names stable from one run to the next.
    """
    used = {r.casefold() for r in reserved}
    result = {}
    for ident, desired, _ in sorted(entries, key=lambda e: e[2]):
        name, n = desired, 2
        while name.casefold() in used:
            name = f"{desired} ({n})"
            n += 1
        used.add(name.casefold())
        result[ident] = name
    return result


def rel_link(target: str, from_dir: str) -> str:
    """Relative, URL-encoded link from a folder to a file (both vault-relative)."""
    return quote(posixpath.relpath(target, from_dir or "."))


# --------------------------------------------------------------------------
# Auth and HTTP
# --------------------------------------------------------------------------

class ExportError(Exception):
    pass


class GraphError(ExportError):
    def __init__(self, status: int, message: str, url: str = ""):
        super().__init__(f"HTTP {status}: {message}")
        self.status = status
        self.url = url


class Auth:
    """Device-code sign-in for a personal Microsoft account, with a token cache."""

    def __init__(self, client_id: str):
        import msal  # imported here so the rest of the module works without it

        self.cache_path = CONFIG_DIR / "token_cache.json"
        self.cache = msal.SerializableTokenCache()
        if self.cache_path.exists():
            self.cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id, authority=AUTHORITY, token_cache=self.cache
        )

    def _save(self):
        if not self.cache.has_state_changed:
            return
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        tmp = self.cache_path.with_suffix(".tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(self.cache.serialize())
        os.replace(tmp, self.cache_path)

    def login(self) -> str:
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise ExportError(
                "Could not start sign-in: "
                + str(flow.get("error_description") or flow)
            )
        print(flow["message"], flush=True)
        result = self.app.acquire_token_by_device_flow(flow)
        self._save()
        if "access_token" not in result:
            raise ExportError(
                "Sign-in failed: " + str(result.get("error_description") or result)
            )
        return (result.get("id_token_claims") or {}).get("preferred_username", "")

    def token(self, force_refresh: bool = False) -> str:
        accounts = self.app.get_accounts()
        if not accounts:
            raise ExportError("Not signed in. Run: onenote_export.py login")
        result = self.app.acquire_token_silent(
            SCOPES, account=accounts[0], force_refresh=force_refresh
        )
        self._save()
        if not result or "access_token" not in result:
            raise ExportError("Sign-in expired. Run: onenote_export.py login")
        return result["access_token"]


class GraphClient:
    """Paced Graph GET requests with retry on throttling and transient errors."""

    RETRY_STATUS = {429, 500, 502, 503, 504}

    def __init__(self, token_provider, delay=0.5, max_tries=8,
                 session=None, sleep=time.sleep, log=None):
        self.token_provider = token_provider
        self.delay = delay
        self.max_tries = max_tries
        self.session = session or requests.Session()
        self.sleep = sleep
        self.log = log or (lambda msg: None)
        self._last = 0.0
        self.request_count = 0

    def _request(self, url, params=None):
        if not url.startswith("http"):
            url = GRAPH + url
        if urlparse(url).hostname != GRAPH_HOST:
            # Never send the access token anywhere but Microsoft Graph.
            raise GraphError(0, "refusing to call a non-Graph host", url)
        refreshed = False
        last_error = "no response"
        for attempt in range(self.max_tries):
            wait = self._last + self.delay - time.monotonic()
            if wait > 0:
                self.sleep(wait)
            headers = {"Authorization": "Bearer " + self.token_provider(False)}
            try:
                self.request_count += 1
                resp = self.session.get(url, params=params, headers=headers, timeout=90)
            except requests.RequestException as exc:
                self._last = time.monotonic()
                last_error = str(exc)
                self._backoff(attempt, None, last_error)
                continue
            self._last = time.monotonic()
            if resp.status_code == 401 and not refreshed:
                refreshed = True
                self.token_provider(True)
                continue
            if resp.status_code in self.RETRY_STATUS:
                last_error = f"HTTP {resp.status_code}"
                self._backoff(attempt, resp.headers.get("Retry-After"), last_error)
                continue
            if resp.status_code >= 400:
                raise GraphError(resp.status_code, _error_message(resp), url)
            return resp
        raise GraphError(0, f"gave up after {self.max_tries} tries ({last_error})", url)

    def _backoff(self, attempt, retry_after, reason):
        try:
            pause = float(retry_after)
        except (TypeError, ValueError):
            pause = min(120.0, 2.0 ** (attempt + 1)) + random.uniform(0, 1)
        self.log(f"  {reason}; waiting {pause:.0f}s before retrying")
        self.sleep(pause)

    def get_all(self, path, params=None):
        """GET a collection, following @odata.nextLink to the end."""
        items, url = [], path
        while url:
            data = self._request(url, params).json()
            items.extend(data.get("value", []))
            url, params = data.get("@odata.nextLink"), None
        return items

    def get_text(self, url):
        resp = self._request(url)
        resp.encoding = "utf-8"
        return resp.text

    def get_bytes(self, url):
        return self._request(url).content


def _error_message(resp) -> str:
    try:
        err = resp.json().get("error", {})
        return f"{err.get('code', '')} {err.get('message', '')}".strip() or resp.reason
    except Exception:
        return (resp.text or resp.reason or "")[:200]


# --------------------------------------------------------------------------
# Crawling the hierarchy
# --------------------------------------------------------------------------

@dataclass
class Page:
    id: str
    title: str
    created: str
    modified: str
    level: int
    order: int
    guid: str
    web_url: str
    notebook_id: str
    parent: "Page | None" = None
    stem: str = ""
    dir: str = ""
    section_dir: str = ""

    @property
    def rel_path(self) -> str:
        return posixpath.join(self.dir, self.stem + ".md")


@dataclass
class Container:
    kind: str  # notebook | group | section
    id: str
    name: str
    created: str
    children: list = field(default_factory=list)
    pages: list = field(default_factory=list)
    failed: bool = False
    dir: str = ""


_GUID_RE = re.compile(r"page-id=\{?([0-9a-fA-F-]{36})\}?")


def _page_from_json(obj, notebook_id) -> Page:
    links = obj.get("links") or {}
    client_url = (links.get("oneNoteClientUrl") or {}).get("href", "")
    match = _GUID_RE.search(client_url)
    return Page(
        id=obj["id"],
        title=(obj.get("title") or "").strip(),
        created=obj.get("createdDateTime") or "",
        modified=obj.get("lastModifiedDateTime") or obj.get("createdDateTime") or "",
        level=int(obj.get("level") or 0),
        order=int(obj.get("order") or 0),
        guid=match.group(1).lower() if match else "",
        web_url=(links.get("oneNoteWebUrl") or {}).get("href", ""),
        notebook_id=notebook_id,
    )


def crawl(client, only=None, log=print):
    """Read notebooks, section groups, sections and page metadata.

    Returns (notebooks, errors). A container that could not be read is kept
    with failed=True so that its pages are never treated as deleted.
    """
    errors = []
    wanted = {n.casefold() for n in only} if only else None
    notebooks = []
    for nb in client.get_all("/me/onenote/notebooks"):
        name = nb.get("displayName") or ""
        if wanted is not None and name.casefold() not in wanted:
            continue
        node = Container("notebook", nb["id"], name, nb.get("createdDateTime") or "")
        log(f"Reading notebook: {name}")
        _fill(client, node, f"/me/onenote/notebooks/{nb['id']}", nb["id"], errors, log)
        notebooks.append(node)
    if wanted is not None:
        found = {n.name.casefold() for n in notebooks}
        for missing in sorted(wanted - found):
            errors.append(f"notebook not found: {missing}")
    return notebooks, errors


def _fill(client, node, base, notebook_id, errors, log):
    try:
        sections = client.get_all(base + "/sections")
        groups = client.get_all(base + "/sectionGroups")
    except GraphError as exc:
        node.failed = True
        errors.append(f"{node.kind} '{node.name}': {exc}")
        return
    for sec in sections:
        child = Container("section", sec["id"], sec.get("displayName") or "",
                          sec.get("createdDateTime") or "")
        node.children.append(child)
        try:
            raw = _list_pages(client, sec["id"])
        except GraphError as exc:
            child.failed = True
            errors.append(f"section '{child.name}': {exc}")
            continue
        child.pages = [_page_from_json(p, notebook_id) for p in raw]
        log(f"  {child.name}: {len(child.pages)} pages")
    for grp in groups:
        child = Container("group", grp["id"], grp.get("displayName") or "",
                          grp.get("createdDateTime") or "")
        node.children.append(child)
        _fill(client, child, f"/me/onenote/sectionGroups/{grp['id']}",
              notebook_id, errors, log)


def _list_pages(client, section_id):
    path = f"/me/onenote/sections/{section_id}/pages"
    try:
        # pagelevel=true adds the level and order needed to nest subpages.
        return client.get_all(path, {"pagelevel": "true", "$top": "100"})
    except GraphError as exc:
        if exc.status != 400:
            raise
        return client.get_all(path, {"$top": "100"})


# --------------------------------------------------------------------------
# Planning local paths
# --------------------------------------------------------------------------

def plan_paths(notebooks):
    """Assign a folder to every container and a file path to every page.

    Returns the flat list of pages, in export order.
    """
    pages = []
    names = unique_names(
        [(n.id, sanitize_name(n.name, "Notebook"), (n.created, n.id)) for n in notebooks],
        reserved=[STATE_DIR],
    )
    for nb in notebooks:
        nb.dir = names[nb.id]
        _plan_container(nb, pages)
    return pages


def _plan_container(node, pages):
    names = unique_names(
        [(c.id, sanitize_name(c.name, "Section"), (c.created, c.id)) for c in node.children]
    )
    for child in node.children:
        child.dir = posixpath.join(node.dir, names[child.id])
        if child.kind == "section":
            _plan_section(child, pages)
        else:
            _plan_container(child, pages)


def _plan_section(section, pages):
    ordered = sorted(section.pages, key=lambda p: p.order)  # stable sort
    last_at_level = {}
    for page in ordered:
        page.parent = None
        for lvl in range(page.level - 1, -1, -1):
            if lvl in last_at_level:
                page.parent = last_at_level[lvl]
                break
        last_at_level[page.level] = page
        for lvl in [k for k in last_at_level if k > page.level]:
            del last_at_level[lvl]

    def place(parent, directory):
        siblings = [p for p in ordered if p.parent is parent]
        names = unique_names(
            [(p.id, _page_name(p), (p.created, p.id)) for p in siblings],
            reserved=[ASSETS_DIR],
        )
        for p in siblings:
            p.stem, p.dir, p.section_dir = names[p.id], directory, section.dir
            pages.append(p)
            place(p, posixpath.join(directory, p.stem))

    place(None, section.dir)


def _page_name(page) -> str:
    fallback = "Untitled " + (page.created[:10] or "page")
    return sanitize_name(page.title, fallback.strip())


# --------------------------------------------------------------------------
# HTML to Markdown
# --------------------------------------------------------------------------

_TODO_OPEN = "ONXTODOOPENX"
_TODO_DONE = "ONXTODODONEX"
_TODO_LINE = re.compile(
    rf"^(\s*)(?:(?:[-*+]|\d+\.)\s+)?({_TODO_OPEN}|{_TODO_DONE})\s*(.*)$"
)
_TODO_ANY = re.compile(rf"({_TODO_OPEN}|{_TODO_DONE})\s*")


def html_to_markdown(html, resolve_resource=None, resolve_link=None) -> str:
    """Convert OneNote page HTML to Markdown.

    resolve_resource(url, mime_type, filename_or_None) -> local link or None
    resolve_link(href) -> replacement href or None
    """
    soup = BeautifulSoup(html, "html.parser")
    body = soup.body or soup

    for img in body.find_all("img"):
        src = img.get("data-fullres-src") or img.get("src")
        mime = img.get("data-fullres-src-type") or img.get("data-src-type")
        if src and resolve_resource:
            local = resolve_resource(src, mime, None)
            if local:
                img["src"] = local
        img["alt"] = " ".join((img.get("alt") or "").split())

    for obj in body.find_all("object"):
        name = obj.get("data-attachment") or "attachment"
        url = obj.get("data") or ""
        local = resolve_resource(url, obj.get("type"), name) if url and resolve_resource else None
        link = soup.new_tag("a", href=local or url)
        link.string = name
        holder = soup.new_tag("p")
        holder.append(link)
        obj.replace_with(holder)

    if resolve_link:
        for a in body.find_all("a", href=True):
            new = resolve_link(a["href"])
            if new:
                a["href"] = new

    # OneNote expresses most formatting as inline styles on spans.
    for span in body.find_all("span", style=True):
        style = span["style"].replace(" ", "").lower()
        wrappers = []
        if "font-weight:bold" in style:
            wrappers.append("b")
        if "font-style:italic" in style:
            wrappers.append("i")
        if "line-through" in style:
            wrappers.append("del")
        if wrappers:
            span.name = wrappers[0]
            del span["style"]
            for tag in wrappers[1:]:
                inner = soup.new_tag(tag)
                for child in list(span.contents):
                    inner.append(child.extract())
                span.append(inner)

    # To-do tags become task-list checkboxes.
    for el in body.find_all(attrs={"data-tag": True}):
        for tag in el["data-tag"].split(","):
            tag = tag.strip().lower()
            if tag.startswith("to-do"):
                el.insert(0, (_TODO_DONE if "completed" in tag else _TODO_OPEN) + " ")
                break

    # Tables with merged cells or nested tables cannot be Markdown tables;
    # keep them as HTML, which Obsidian and most viewers render.
    raw_blocks = []
    for table in body.find_all("table"):
        if table.find_parent("table") is not None:
            continue
        if table.find("table") is None and not _has_merged_cells(table):
            continue
        for el in [table, *table.find_all(True)]:
            for attr in ("style", "data-id", "id"):
                el.attrs.pop(attr, None)
        token = f"ONXRAWHTML{len(raw_blocks)}X"
        raw_blocks.append(str(table))
        holder = soup.new_tag("p")
        holder.string = token
        table.replace_with(holder)

    text = markdownify(body.decode_contents(), heading_style="ATX", bullets="-")

    lines = []
    for line in text.splitlines():
        match = _TODO_LINE.match(line)
        if match:
            indent, marker, rest = match.groups()
            box = "x" if marker == _TODO_DONE else " "
            line = f"{indent}- [{box}] {_TODO_ANY.sub('', rest)}"
        elif _TODO_OPEN in line or _TODO_DONE in line:
            line = line.replace(_TODO_OPEN, "☐").replace(_TODO_DONE, "☑")
        lines.append(line.rstrip())
    text = "\n".join(lines)

    # Consecutive checkbox paragraphs read better as one tight list.
    text = re.sub(r"(?m)^(\s*- \[[ x]\] .*)\n\n+(?=\s*- \[[ x]\] )", r"\1\n", text)
    for i, raw in enumerate(raw_blocks):
        text = text.replace(f"ONXRAWHTML{i}X", f"\n\n{raw}\n\n")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text + "\n" if text else ""


def _has_merged_cells(table) -> bool:
    for cell in table.find_all(["td", "th"]):
        for attr in ("colspan", "rowspan"):
            try:
                if int(cell.get(attr) or 1) > 1:
                    return True
            except ValueError:
                pass
    return False


# --------------------------------------------------------------------------
# Writing files
# --------------------------------------------------------------------------

def _write_atomic(path: Path, data, mode="w"):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    kwargs = {"encoding": "utf-8", "newline": "\n"} if mode == "w" else {}
    with open(tmp, mode, **kwargs) as f:
        f.write(data)
    os.replace(tmp, path)


def _set_mtime(path: Path, iso: str):
    try:
        stamp = datetime.fromisoformat(re.sub(r"(\.\d{6})\d+", r"\1", iso).replace("Z", "+00:00"))
        os.utime(path, (stamp.timestamp(), stamp.timestamp()))
    except (ValueError, OSError):
        pass


def _front_matter(page: Page, metadata: bool = False) -> str:
    def q(value):  # a JSON string is also a valid YAML double-quoted string
        return json.dumps(value, ensure_ascii=False)

    rows = []
    effective_title = page.title or page.stem
    if unicodedata.normalize("NFC", effective_title) != page.stem:
        rows.append(f"title: {q(effective_title)}")
    if metadata:
        rows.extend([
            f"created: {page.created}",
            f"modified: {page.modified}",
            "source: onenote",
            f"onenote_id: {q(page.id)}",
        ])
        if page.web_url:
            rows.append(f"onenote_url: {q(page.web_url)}")
    if not rows:
        return ""
    return "---\n" + "\n".join(rows) + "\n---\n\n"


_EXT_FIXES = {"image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png",
              "image/gif": ".gif", "image/bmp": ".bmp", "image/svg+xml": ".svg",
              "image/tiff": ".tiff", "image/webp": ".webp"}


def _extension(mime) -> str:
    mime = (mime or "").split(";")[0].strip().lower()
    return _EXT_FIXES.get(mime) or mimetypes.guess_extension(mime) or ".bin"


def export_page(client, out_dir: Path, page: Page, guid_map, warn, metadata: bool = False):
    """Fetch one page, save its images and attachments, write the Markdown file."""
    html = client.get_text(f"/me/onenote/pages/{page.id}/content")
    assets_dir = posixpath.join(page.section_dir, ASSETS_DIR)

    def resolve_resource(url, mime, filename):
        if urlparse(url).hostname != GRAPH_HOST:
            return None  # an ordinary web image or link: leave it alone
        match = re.search(r"/resources/([^/]+)/", url)
        digest = hashlib.sha1((match.group(1) if match else url).encode()).hexdigest()
        if filename:
            name = f"{digest[:8]}-{sanitize_name(filename, 'attachment')}"
        else:
            name = digest[:16] + _extension(mime)
        rel = posixpath.join(assets_dir, name)
        target = out_dir / rel
        if not target.exists():
            try:
                _write_atomic(target, client.get_bytes(url), "wb")
            except GraphError as exc:
                warn(f"{page.rel_path}: could not download {filename or 'image'} ({exc})")
                return None
        return rel_link(rel, page.dir)

    def resolve_link(href):
        match = _GUID_RE.search(href)
        if not match:
            return None
        target = guid_map.get(match.group(1).lower())
        if not target or target == page.rel_path:
            return None
        return rel_link(target, page.dir)

    body = html_to_markdown(html, resolve_resource, resolve_link)
    front = _front_matter(page, metadata=metadata)
    if not front and body.startswith("---"):
        body = "\n" + body  # otherwise a leading horizontal rule reads as front matter
    path = out_dir / page.rel_path
    _write_atomic(path, front + body)
    _set_mtime(path, page.modified)


# --------------------------------------------------------------------------
# Sync
# --------------------------------------------------------------------------

def load_manifest(out_dir: Path) -> dict:
    path = out_dir / STATE_DIR / "manifest.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {"version": 1, "pages": {}}


def save_manifest(out_dir: Path, manifest: dict):
    _write_atomic(out_dir / STATE_DIR / "manifest.json",
                  json.dumps(manifest, ensure_ascii=False, indent=1))


def _remove_file(out_dir: Path, rel: str):
    path = out_dir / rel
    try:
        path.unlink()
    except FileNotFoundError:
        return
    parent = path.parent
    while parent != out_dir:
        try:
            parent.rmdir()  # only succeeds when empty
        except OSError:
            break
        parent = parent.parent


def _removable(out_dir: Path, old: str, new: str, owners: dict, page_id: str) -> bool:
    """May the file at `old` be deleted after page `page_id` moved to `new`?

    Not if another page now has that path (a name freed by one page is often
    taken by another), and not if it is the same file as `new` (a case-only
    rename on a case-insensitive file system).
    """
    if owners.get(old.casefold(), page_id) != page_id:
        return False
    try:
        if os.path.samefile(out_dir / old, out_dir / new):
            return False
    except OSError:
        pass
    return True


def _all_containers(nodes):
    for node in nodes:
        yield node
        yield from _all_containers(node.children)


def sync(client, out_dir: Path, only=None, force=False, prune=False, log=print,
         metadata: bool = False):
    """Export new and changed pages. Returns a dict of counts and messages."""
    out_dir.mkdir(parents=True, exist_ok=True)
    notebooks, errors = crawl(client, only, log)
    pages = plan_paths(notebooks)
    guid_map = {p.guid: p.rel_path for p in pages if p.guid}
    owners = {p.rel_path.casefold(): p.id for p in pages}
    manifest = load_manifest(out_dir)
    known = manifest["pages"]
    warnings = []
    stats = {"exported": 0, "unchanged": 0, "failed": 0, "removed": 0, "stale": 0}

    todo = []
    for page in pages:
        entry = known.get(page.id)
        if (not force and entry and entry.get("modified") == page.modified
                and entry.get("path") == page.rel_path
                and bool(entry.get("metadata", True)) == bool(metadata)
                and (out_dir / page.rel_path).exists()):
            stats["unchanged"] += 1
        else:
            todo.append(page)

    log(f"{len(pages)} pages found; {len(todo)} to export, {stats['unchanged']} up to date.")
    for i, page in enumerate(todo, 1):
        log(f"[{i}/{len(todo)}] {page.rel_path}")
        try:
            export_page(client, out_dir, page, guid_map, warnings.append, metadata=metadata)
        except (GraphError, OSError) as exc:
            stats["failed"] += 1
            errors.append(f"page '{page.rel_path}': {exc}")
            continue
        old = known.get(page.id, {}).get("path")
        if old and old != page.rel_path and _removable(
                out_dir, old, page.rel_path, owners, page.id):
            _remove_file(out_dir, old)  # the page was renamed or moved
        known[page.id] = {"path": page.rel_path, "modified": page.modified,
                          "notebook": page.notebook_id, "metadata": bool(metadata)}
        save_manifest(out_dir, manifest)  # after every page, so a rerun resumes
        stats["exported"] += 1

    # Pages that are gone from OneNote. Skipped when anything failed to list,
    # so that a temporary error can never delete local files.
    crawled = {nb.id for nb in notebooks}
    listing_ok = not any(c.failed for c in _all_containers(notebooks))
    current = {p.id for p in pages}
    gone = [pid for pid, e in known.items()
            if e.get("notebook") in crawled and pid not in current]
    if gone and prune and listing_ok:
        for pid in gone:
            if known[pid]["path"].casefold() not in owners:
                _remove_file(out_dir, known[pid]["path"])
            del known[pid]
        stats["removed"] = len(gone)
    else:
        stats["stale"] = len(gone)
    save_manifest(out_dir, manifest)

    stats["errors"], stats["warnings"] = errors, warnings
    return stats


def print_tree(notebooks, log=print):
    def show(node, depth):
        label = {"notebook": "Notebook", "group": "Group", "section": "Section"}[node.kind]
        note = "  (could not be read)" if node.failed else ""
        log(f"{'  ' * depth}{label}: {node.name}  ->  {node.dir}/{note}")
        if node.kind == "section":
            for page in sorted(node.pages, key=lambda p: p.rel_path):
                extra = posixpath.relpath(page.rel_path, node.dir)
                log(f"{'  ' * (depth + 1)}{extra}")
        for child in node.children:
            show(child, depth + 1)

    for nb in notebooks:
        show(nb, 0)


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def _load_config() -> dict:
    path = CONFIG_DIR / "config.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def _client_id(args) -> str:
    cid = args.client_id or os.environ.get("ONENOTE_CLIENT_ID") or _load_config().get("client_id")
    if not cid:
        raise ExportError(
            "No client ID. Run: onenote_export.py login --client-id <Application (client) ID>"
        )
    return cid


def _make_client(args) -> GraphClient:
    auth = Auth(_client_id(args))
    return GraphClient(auth.token, delay=args.delay, log=print)


def cmd_login(args):
    cid = _client_id(args)
    name = Auth(cid).login()
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    config = _load_config()
    config["client_id"] = cid
    _write_atomic(CONFIG_DIR / "config.json", json.dumps(config, indent=1))
    print(f"Signed in{' as ' + name if name else ''}. Token cached in {CONFIG_DIR}.")
    return 0


def cmd_list(args):
    client = _make_client(args)
    notebooks, errors = crawl(client, args.notebook, log=lambda msg: None)
    pages = plan_paths(notebooks)
    print_tree(notebooks)
    print(f"\n{len(notebooks)} notebooks, {len(pages)} pages.")
    for err in errors:
        print("ERROR:", err, file=sys.stderr)
    return 1 if errors else 0


def cmd_sync(args):
    client = _make_client(args)
    metadata = args.onenote_metadata
    if metadata is None:  # not given on the command line: use config.json, else off
        metadata = bool(_load_config().get("onenote_metadata", False))
    stats = sync(client, Path(args.out).expanduser(), only=args.notebook,
                 force=args.force, prune=args.prune, metadata=metadata)
    print(f"\nDone: {stats['exported']} exported, {stats['unchanged']} unchanged, "
          f"{stats['failed']} failed, {stats['removed']} removed. "
          f"{client.request_count} requests.")
    if stats["stale"]:
        print(f"{stats['stale']} local pages are no longer in OneNote and were kept. "
              "Rerun with --prune to delete them.")
    for msg in stats["warnings"]:
        print("WARNING:", msg, file=sys.stderr)
    for msg in stats["errors"]:
        print("ERROR:", msg, file=sys.stderr)
    return 1 if stats["errors"] else 0


def build_parser():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--client-id", help="Application (client) ID of your app registration")
    common.add_argument("--delay", type=float, default=0.5,
                        help="minimum seconds between requests (default 0.5)")
    parser = argparse.ArgumentParser(
        prog="onenote_export.py",
        description="Export OneNote notebooks to a local Markdown tree.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("login", parents=[common], help="sign in with a device code")
    p.set_defaults(func=cmd_login)

    p = sub.add_parser("list", parents=[common], help="show the hierarchy and planned paths")
    p.add_argument("--notebook", action="append", help="only this notebook (repeatable)")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("sync", parents=[common], help="export new and changed pages")
    p.add_argument("--out", required=True, help="folder to write the Markdown tree into")
    p.add_argument("--notebook", action="append", help="only this notebook (repeatable)")
    p.add_argument("--force", action="store_true", help="re-export every page")
    p.add_argument("--prune", action="store_true",
                   help="delete local files for pages removed from OneNote")
    p.add_argument("--onenote-metadata", action=argparse.BooleanOptionalAction, default=None,
                   help="add OneNote fields (created, modified, source, onenote_id, onenote_url) "
                        "to each page's front matter; off by default. "
                        "Use --no-onenote-metadata to override the config file")
    p.set_defaults(func=cmd_sync)
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except ExportError as exc:
        print("ERROR:", exc, file=sys.stderr)
        return 2
    except requests.RequestException as exc:
        print("ERROR: could not reach Microsoft:", exc, file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("\nInterrupted. Run the same command again to resume.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
