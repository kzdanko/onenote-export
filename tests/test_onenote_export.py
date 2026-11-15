"""Offline tests: no Microsoft account or network needed.

Run from the project folder:  python3 -m unittest discover -s tests -v
"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import onenote_export as oe  # noqa: E402

RES = "https://graph.microsoft.com/v1.0/users('me')/onenote/resources"


def page_html(body):
    return ("<html lang='en-US'><head><title>t</title></head>"
            "<body data-absolute-enabled='true' style='font-family:Calibri'>"
            f"<div style='position:absolute;left:48px;top:115px;width:624px'>{body}</div>"
            "</body></html>")


class FakeClient:
    """Stands in for GraphClient, serving a small made-up account."""

    def __init__(self):
        self.collections = {}
        self.content = {}
        self.resources = {}
        self.fetched_pages = []
        self.fail = set()
        self.request_count = 0

    def get_all(self, path, params=None):
        if path in self.fail:
            raise oe.GraphError(403, "forbidden", path)
        return list(self.collections.get(path, []))

    def get_text(self, url):
        pid = url.split("/pages/")[1].split("/")[0]
        self.fetched_pages.append(pid)
        return self.content[pid]

    def get_bytes(self, url):
        return self.resources[url]


def guid(n):
    return f"{n:08d}-0000-0000-0000-000000000000"


def page(pid, title, order, level=0, modified="2026-01-02T03:04:05Z", n=None):
    n = n if n is not None else order + 1
    return {
        "id": pid, "title": title, "level": level, "order": order,
        "createdDateTime": f"2025-01-{n:02d}T00:00:00Z",
        "lastModifiedDateTime": modified,
        "links": {
            "oneNoteClientUrl": {"href": "onenote:https://d.docs.live.net/x/Nb/S.one#"
                                         f"{title}&section-id={{{guid(99)}}}&page-id={{{guid(n)}}}&end"},
            "oneNoteWebUrl": {"href": f"https://onedrive.live.com/view?id={pid}"},
        },
    }


def make_account():
    c = FakeClient()
    c.collections["/me/onenote/notebooks"] = [
        {"id": "nb1", "displayName": "Jegyzetek: Árvíztűrő", "createdDateTime": "2020-01-01T00:00:00Z"},
        {"id": "nb2", "displayName": "Work", "createdDateTime": "2021-01-01T00:00:00Z"},
    ]
    c.collections["/me/onenote/notebooks/nb1/sections"] = [
        {"id": "s1", "displayName": "Napló", "createdDateTime": "2020-01-02T00:00:00Z"}]
    c.collections["/me/onenote/notebooks/nb1/sectionGroups"] = [
        {"id": "g1", "displayName": "Archive", "createdDateTime": "2020-01-03T00:00:00Z"}]
    c.collections["/me/onenote/sectionGroups/g1/sections"] = [
        {"id": "s2", "displayName": "Old/Stuff", "createdDateTime": "2020-01-04T00:00:00Z"}]
    c.collections["/me/onenote/notebooks/nb2/sections"] = [
        {"id": "s3", "displayName": "Meetings", "createdDateTime": "2021-01-02T00:00:00Z"}]
    c.collections["/me/onenote/sections/s1/pages"] = [
        page("p1", "Első bejegyzés", 0),
        page("p2", "Child page", 1, level=1),
        page("p3", "Grandchild", 2, level=2),
        page("p4", "Ideas", 3),
        page("p5", "ideas", 4),          # same name ignoring case
        page("p6", "", 5),               # untitled
    ]
    c.collections["/me/onenote/sections/s2/pages"] = [page("p7", "Old note?", 0, n=20)]
    c.collections["/me/onenote/sections/s3/pages"] = [page("p8", "Standup", 0, n=30)]
    img = f"{RES}/0-abc!1-XYZ/$value"
    pdf = f"{RES}/0-def!1-XYZ/$value"
    c.resources = {img: b"\x89PNG fake", pdf: b"%PDF fake"}
    c.content = {
        "p1": page_html(
            "<h1>Heading</h1><p>Árvíztűrő <span style='font-weight:bold'>tükörfúrógép</span></p>"
            f'<img src="{img}" data-src-type="image/png" data-fullres-src="{img}" '
            "data-fullres-src-type='image/png' alt='a picture' />"
            f'<object data-attachment="report 1.pdf" type="application/pdf" data="{pdf}" />'
            "<p><a href='onenote:https://d.docs.live.net/x/Nb/S.one#Ideas&amp;"
            f"section-id={{{guid(99)}}}&amp;page-id={{{guid(4)}}}&amp;end'>see ideas</a></p>"),
        "p2": page_html(f'<p>child</p><img src="{img}" data-src-type="image/png" />'),
        "p3": page_html("<p>grandchild</p>"),
        "p4": page_html("<p data-tag='to-do'>buy milk</p><p data-tag='to-do:completed'>call</p>"),
        "p5": page_html("<p>lower</p>"),
        "p6": page_html("<p>no title</p>"),
        "p7": page_html("<p>old</p>"),
        "p8": page_html("<p>standup</p>"),
    }
    return c


class SanitizeTests(unittest.TestCase):
    def test_keeps_accents_and_removes_forbidden(self):
        self.assertEqual(oe.sanitize_name("Árvíztűrő: tükör/fúró?"), "Árvíztűrő- tükör-fúró")

    def test_edges(self):
        self.assertEqual(oe.sanitize_name("  trailing dots... "), "trailing dots")
        self.assertEqual(oe.sanitize_name("???"), "Untitled")
        self.assertEqual(oe.sanitize_name("CON"), "_CON")
        self.assertEqual(oe.sanitize_name("a [b] #c"), "a b c")
        self.assertLessEqual(len(oe.sanitize_name("x" * 500)), 120)

    def test_unique_names_is_case_insensitive_and_stable(self):
        names = oe.unique_names([("b", "Ideas", ("2025", "b")), ("a", "ideas", ("2024", "a"))])
        self.assertEqual(names, {"a": "ideas", "b": "Ideas (2)"})


class ConvertTests(unittest.TestCase):
    def test_basic_formatting(self):
        md = oe.html_to_markdown(page_html(
            "<h2>Title</h2><p>plain <span style='font-weight:bold;font-style:italic'>both</span> "
            "<span style='text-decoration:line-through'>gone</span></p>"
            "<ul><li>one</li><li>two</li></ul>"))
        self.assertIn("## Title", md)
        self.assertIn("***both***", md)
        self.assertIn("~~gone~~", md)
        self.assertIn("- one\n- two", md)
        self.assertNotIn("position", md)

    def test_todo_tags(self):
        md = oe.html_to_markdown(page_html(
            "<p data-tag='to-do'>first</p><p data-tag='to-do:completed,important'>second</p>"
            "<ul><li data-tag='to-do'>in list</li></ul><p>after</p>"))
        self.assertIn("- [ ] first\n- [x] second", md)
        self.assertIn("- [ ] in list", md)
        self.assertNotIn("ONX", md)

    def test_simple_table_is_markdown(self):
        md = oe.html_to_markdown(page_html(
            "<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>"))
        self.assertIn("| a | b |", md)
        self.assertIn("| 1 | 2 |", md)
        self.assertNotIn("<table", md)

    def test_merged_table_stays_html(self):
        md = oe.html_to_markdown(page_html(
            "<p>before</p><table style='border:1px'><tr><td colspan='2'>wide</td></tr>"
            "<tr><td>1</td><td>2</td></tr></table><p>after</p>"))
        self.assertIn('<td colspan="2">wide</td>', md)
        self.assertNotIn("style=", md)
        self.assertNotIn("ONX", md)
        self.assertLess(md.index("before"), md.index("<table"))
        self.assertLess(md.index("</table>"), md.index("after"))

    def test_external_image_untouched(self):
        md = oe.html_to_markdown(
            page_html("<img src='https://example.com/a.png' alt='x'/>"),
            resolve_resource=lambda url, mime, name: None)
        self.assertIn("![x](https://example.com/a.png)", md)


class SyncTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "vault"
        self.client = make_account()
        self.quiet = lambda msg: None

    def tearDown(self):
        self.tmp.cleanup()

    def run_sync(self, **kw):
        return oe.sync(self.client, self.out, log=self.quiet, **kw)

    def files(self):
        return sorted(str(p.relative_to(self.out)).replace("\\", "/")
                      for p in self.out.rglob("*") if p.is_file() and ".onenote-export" not in p.parts)

    def test_full_export_layout(self):
        stats = self.run_sync()
        self.assertEqual(stats["errors"], [])
        self.assertEqual(stats["exported"], 8)
        nb = "Jegyzetek- Árvíztűrő"
        md = [f for f in self.files() if f.endswith(".md")]
        self.assertEqual(md, sorted([
            f"{nb}/Napló/Első bejegyzés.md",
            f"{nb}/Napló/Első bejegyzés/Child page.md",
            f"{nb}/Napló/Első bejegyzés/Child page/Grandchild.md",
            f"{nb}/Napló/Ideas.md",
            f"{nb}/Napló/ideas (2).md",
            f"{nb}/Napló/Untitled 2025-01-06.md",
            f"{nb}/Archive/Old-Stuff/Old note.md",
            "Work/Meetings/Standup.md",
        ]))
        assets = [f for f in self.files() if "/_assets/" in f]
        self.assertEqual(len(assets), 2)  # the image is shared by two pages, saved once
        self.assertTrue(all(f.startswith(f"{nb}/Napló/_assets/") for f in assets))

    def test_page_content_links_and_front_matter(self):
        self.run_sync(metadata=True)
        base = self.out / "Jegyzetek- Árvíztűrő" / "Napló"
        text = (base / "Első bejegyzés.md").read_text(encoding="utf-8")
        # When title equals filename, title frontmatter is omitted
        self.assertTrue(text.startswith("---\ncreated: 2025-01-01"))
        self.assertNotIn('title:', text)
        self.assertIn('onenote_id: "p1"', text)
        self.assertIn("source: onenote", text)
        # When title differs from filename, title frontmatter is preserved
        old_note = (self.out / "Jegyzetek- Árvíztűrő" / "Archive" / "Old-Stuff" / "Old note.md").read_text(encoding="utf-8")
        self.assertIn('title: "Old note?"', old_note)
        self.assertIn("Árvíztűrő **tükörfúrógép**", text)
        self.assertRegex(text, r"!\[a picture\]\(_assets/[0-9a-f]{16}\.png\)")
        self.assertRegex(text, r"\[report 1\.pdf\]\(_assets/[0-9a-f]{8}-report%201\.pdf\)")
        self.assertIn("[see ideas](Ideas.md)", text)
        child = (base / "Első bejegyzés" / "Child page.md").read_text(encoding="utf-8")
        self.assertRegex(child, r"!\[\]\(\.\./_assets/[0-9a-f]{16}\.png\)")
        # every relative link in every file resolves to a real file
        import re
        from urllib.parse import unquote
        for f in self.out.rglob("*.md"):
            for target in re.findall(r"\]\(([^)]+)\)", f.read_text(encoding="utf-8")):
                if "://" not in target:
                    self.assertTrue((f.parent / unquote(target)).exists(), f"{f}: {target}")

    def test_default_front_matter_has_no_onenote_metadata(self):
        self.run_sync()
        base = self.out / "Jegyzetek- Árvíztűrő" / "Napló"
        text = (base / "Első bejegyzés.md").read_text(encoding="utf-8")
        # title == file name: nothing to say, so no front matter at all
        self.assertFalse(text.startswith("---"))
        self.assertTrue(text.startswith("# Heading\n\nÁrvíztűrő **tükörfúrógép**"))
        # title != file name: only the title is kept
        old_note = (self.out / "Jegyzetek- Árvíztűrő" / "Archive" / "Old-Stuff" / "Old note.md").read_text(encoding="utf-8")
        self.assertTrue(old_note.startswith('---\ntitle: "Old note?"\n---\n\n'))
        for field in ("created:", "modified:", "source: onenote", "onenote_id:", "onenote_url:"):
            self.assertNotIn(field, old_note)

    def test_incremental_and_resume(self):
        self.run_sync()
        self.client.fetched_pages.clear()
        stats = self.run_sync()
        self.assertEqual((stats["exported"], stats["unchanged"]), (0, 8))
        self.assertEqual(self.client.fetched_pages, [])

        pages = self.client.collections["/me/onenote/sections/s3/pages"]
        pages[0]["lastModifiedDateTime"] = "2026-06-01T00:00:00Z"
        self.client.content["p8"] = page_html("<p>updated</p>")
        stats = self.run_sync()
        self.assertEqual(self.client.fetched_pages, ["p8"])
        self.assertIn("updated", (self.out / "Work/Meetings/Standup.md").read_text())

        # a file deleted locally is fetched again
        (self.out / "Work/Meetings/Standup.md").unlink()
        self.client.fetched_pages.clear()
        self.run_sync()
        self.assertEqual(self.client.fetched_pages, ["p8"])

    def test_changing_metadata_setting_reexports(self):
        self.run_sync(metadata=True)
        self.client.fetched_pages.clear()
        stats = self.run_sync()  # default: no metadata
        self.assertEqual(stats["exported"], 8)
        standup = (self.out / "Work/Meetings/Standup.md").read_text(encoding="utf-8")
        self.assertNotIn("onenote_id", standup)
        self.assertEqual(self.run_sync()["exported"], 0)
        # a run limited to one notebook must not change the others
        stats = self.run_sync(only=["work"], metadata=True)
        self.assertEqual(stats["exported"], 1)
        other = self.out / "Jegyzetek- Árvíztűrő/Archive/Old-Stuff/Old note.md"
        self.assertNotIn("onenote_id", other.read_text(encoding="utf-8"))
        self.assertEqual(self.run_sync()["exported"], 1)  # only the Work page

    def test_export_made_before_the_setting_existed_counts_as_having_metadata(self):
        self.run_sync(metadata=True)
        path = self.out / ".onenote-export" / "manifest.json"
        manifest = json.loads(path.read_text(encoding="utf-8"))
        for entry in manifest["pages"].values():
            del entry["metadata"]
        path.write_text(json.dumps(manifest), encoding="utf-8")
        self.assertEqual(self.run_sync(metadata=True)["exported"], 0)
        self.assertEqual(self.run_sync()["exported"], 8)

    def test_rename_that_frees_a_name_for_another_page(self):
        pages = self.client.collections["/me/onenote/sections/s1/pages"]
        for p in pages:
            if p["id"] in ("p4", "p5"):
                p["title"] = "Notes"
        self.run_sync()
        base = self.out / "Jegyzetek- Árvíztűrő" / "Napló"
        self.assertTrue((base / "Notes.md").exists() and (base / "Notes (2).md").exists())
        # Rename the page that holds "Notes", and list the other page first, so the
        # page that inherits "Notes.md" is written before the old file is cleaned up.
        for p in pages:
            if p["id"] == "p4":
                p.update(title="Plans", order=4, lastModifiedDateTime="2026-06-01T00:00:00Z")
            if p["id"] == "p5":
                p["order"] = 3
        stats = self.run_sync()
        self.assertEqual(stats["errors"], [])
        self.assertIn("lower", (base / "Notes.md").read_text(encoding="utf-8"))
        self.assertTrue((base / "Plans.md").exists())
        self.assertFalse((base / "Notes (2).md").exists())
        self.assertEqual(self.run_sync()["exported"], 0)  # nothing left to repair

    def test_prune_keeps_a_path_taken_over_by_another_page(self):
        pages = self.client.collections["/me/onenote/sections/s1/pages"]
        for p in pages:
            if p["id"] in ("p4", "p5"):
                p["title"] = "Notes"
        self.run_sync()
        base = self.out / "Jegyzetek- Árvíztűrő" / "Napló"
        # p4 (which owns "Notes.md") is deleted; p5 inherits the name
        pages[:] = [p for p in pages if p["id"] != "p4"]
        stats = self.run_sync(prune=True)
        self.assertEqual(stats["removed"], 1)
        self.assertIn("lower", (base / "Notes.md").read_text(encoding="utf-8"))
        self.assertEqual(self.run_sync()["exported"], 0)

    def test_leading_rule_is_not_front_matter(self):
        self.client.content["p8"] = page_html("<hr/><p>text</p><hr/><p>more</p>")
        self.run_sync()
        text = (self.out / "Work/Meetings/Standup.md").read_text(encoding="utf-8")
        self.assertFalse(text.startswith("---"))
        self.assertIn("text", text)

    def test_decomposed_title_does_not_repeat_the_file_name(self):
        pg = oe.Page("i", "e\u0301lo\u0301", "", "", 0, 0, "", "", "nb")
        pg.stem = oe.sanitize_name(pg.title)
        self.assertEqual(oe._front_matter(pg), "")

    def test_metadata_flag_parsing(self):
        parse = oe.build_parser().parse_args
        self.assertIsNone(parse(["sync", "--out", "x"]).onenote_metadata)
        self.assertTrue(parse(["sync", "--out", "x", "--onenote-metadata"]).onenote_metadata)
        self.assertFalse(parse(["sync", "--out", "x", "--no-onenote-metadata"]).onenote_metadata)

    def test_rename_moves_file(self):
        self.run_sync()
        pages = self.client.collections["/me/onenote/sections/s3/pages"]
        pages[0]["title"] = "Daily standup"
        pages[0]["lastModifiedDateTime"] = "2026-06-02T00:00:00Z"
        self.run_sync()
        self.assertFalse((self.out / "Work/Meetings/Standup.md").exists())
        self.assertTrue((self.out / "Work/Meetings/Daily standup.md").exists())

    def test_deleted_pages_kept_unless_prune(self):
        self.run_sync()
        self.client.collections["/me/onenote/sections/s3/pages"] = []
        stats = self.run_sync()
        self.assertEqual(stats["stale"], 1)
        self.assertTrue((self.out / "Work/Meetings/Standup.md").exists())
        stats = self.run_sync(prune=True)
        self.assertEqual(stats["removed"], 1)
        self.assertFalse((self.out / "Work").exists())

    def test_failed_section_never_prunes(self):
        self.run_sync()
        self.client.fail.add("/me/onenote/sections/s3/pages")
        stats = self.run_sync(prune=True)
        self.assertEqual(stats["removed"], 0)
        self.assertEqual(len(stats["errors"]), 1)
        self.assertTrue((self.out / "Work/Meetings/Standup.md").exists())

    def test_notebook_filter_leaves_other_notebooks_alone(self):
        self.run_sync()
        stats = self.run_sync(only=["work"], prune=True)
        self.assertEqual(stats["removed"], 0)
        self.assertEqual(stats["unchanged"], 1)
        manifest = json.loads((self.out / ".onenote-export/manifest.json").read_text())
        self.assertEqual(len(manifest["pages"]), 8)


class FakeResponse:
    def __init__(self, status, payload=None, headers=None):
        self.status_code, self._payload, self.headers = status, payload or {}, headers or {}
        self.reason, self.text = "reason", ""

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, responses):
        self.responses, self.calls = list(responses), []

    def get(self, url, params=None, headers=None, timeout=None):
        self.calls.append((url, headers["Authorization"]))
        return self.responses.pop(0)


class GraphClientTests(unittest.TestCase):
    def make(self, responses):
        self.sleeps, self.forced = [], []
        session = FakeSession(responses)

        def token(force):
            self.forced.append(force)
            return "tok2" if force or True in self.forced else "tok1"

        client = oe.GraphClient(token, delay=0, session=session, sleep=self.sleeps.append)
        return client, session

    def test_throttle_honours_retry_after_and_paginates(self):
        client, session = self.make([
            FakeResponse(429, headers={"Retry-After": "7"}),
            FakeResponse(200, {"value": [1], "@odata.nextLink": oe.GRAPH + "/next"}),
            FakeResponse(200, {"value": [2]}),
        ])
        self.assertEqual(client.get_all("/me/onenote/notebooks"), [1, 2])
        self.assertEqual(self.sleeps, [7.0])
        self.assertEqual(session.calls[2][0], oe.GRAPH + "/next")

    def test_401_refreshes_token_once(self):
        client, session = self.make([FakeResponse(401), FakeResponse(200, {"value": []})])
        client.get_all("/me/onenote/notebooks")
        self.assertEqual([c[1] for c in session.calls], ["Bearer tok1", "Bearer tok2"])

    def test_error_and_foreign_host(self):
        client, _ = self.make([FakeResponse(404, {"error": {"code": "x", "message": "gone"}})])
        with self.assertRaises(oe.GraphError) as ctx:
            client.get_all("/me/onenote/notebooks")
        self.assertEqual(ctx.exception.status, 404)
        with self.assertRaises(oe.GraphError):
            client.get_bytes("https://evil.example.com/resources/1/$value")


if __name__ == "__main__":
    unittest.main()
