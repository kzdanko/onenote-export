# onenote-export

Copies the OneNote notebooks of a personal Microsoft account into a local
Markdown tree that mirrors the OneNote hierarchy:

```
<out>/
  Notebook/
    Section Group/
      Section/
        Page.md
        Page/              <- subpages of "Page"
          Subpage.md
        _assets/           <- images and attached files for this section
  .onenote-export/manifest.json   <- sync state; do not edit
```

A page gets YAML front matter with its `title` only when the title differs
from the file name, so most files have none; see [Front matter](#front-matter)
for the optional OneNote fields. Rerunning only fetches pages that changed.

## Setup (about 15 minutes)

### 1. Check that the API works for your account

Open <https://developer.microsoft.com/graph/graph-explorer>, sign in with your
personal Microsoft account, and run:

```
GET https://graph.microsoft.com/v1.0/me/onenote/notebooks
```

You may have to grant `Notes.Read` on the "Modify permissions" tab first.
If your notebooks come back, continue. If you get a 401 that will not go
away, stop here: the Graph route is not usable for your account and the tool
needs a different backend (the OneNote desktop COM API).

### 2. Register an app (free)

The tool signs in as you, so Microsoft needs an app ID for it.

1. Go to <https://entra.microsoft.com> (or <https://portal.azure.com>) and
   open **App registrations** > **New registration**. If the portal says your
   account has no directory, sign up for a free Azure account first; that
   creates one.
2. Name: anything, e.g. `onenote-export`.
3. Supported account types: **Personal Microsoft accounts only**.
4. Leave the redirect URI empty and register.
5. In the app, open **Authentication** and set **Allow public client flows**
   to **Yes**. Save.
6. Open **API permissions** > **Add a permission** > **Microsoft Graph** >
   **Delegated permissions** and add `Notes.Read`.
7. Copy the **Application (client) ID** from the Overview page.

Portal labels change from time to time; the settings that matter are personal
accounts, public client flows on, and delegated `Notes.Read`.

### 3. Install (WSL, macOS or Windows; Python 3.9+)

```bash
cd onenote-export
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

### 4. Sign in

```bash
python3 onenote_export.py login --client-id <your client ID>
```

It prints a short code and a URL. Open the URL in any browser, enter the code
and approve. The client ID and token are stored in `~/.config/onenote-export/`
(the token file is readable only by you), so you sign in once.

## Use

```bash
# Dry run: show the hierarchy and the paths that would be written
python3 onenote_export.py list

# Try one notebook first
python3 onenote_export.py sync --out /mnt/c/Users/<you>/OneNoteExport --notebook "My Notebook"

# Everything
python3 onenote_export.py sync --out /mnt/c/Users/<you>/OneNoteExport
```

Options for `sync`:

| Option | Effect |
| --- | --- |
| `--notebook NAME` | Only this notebook. Repeatable. |
| `--force` | Re-export every page, not only changed ones. |
| `--prune` | Delete local files for pages deleted in OneNote. Without it they are kept and counted. |
| `--onenote-metadata` | Also write the OneNote fields (`created`, `modified`, `source`, `onenote_id`, `onenote_url`) into each page's front matter. Off by default. `--no-onenote-metadata` overrides the config file. |
| `--delay SECONDS` | Minimum gap between requests. Default 0.5. Raise it if you see many retries. |

The first run is slow by design: Microsoft throttles the OneNote API, and the
tool waits and retries when told to. You can stop with Ctrl+C and run the same
command again; it resumes where it left off.

## Front matter

By default the files are plain notes, which suits Obsidian: it shows the file
name as the title and keeps its own modified date. A `title:` line appears only
when the OneNote title could not be used as the file name (for example it
contains `/` or `?`, or was shortened or numbered because of a clash).

To record where each note came from, add the OneNote fields:

```bash
python3 onenote_export.py sync --out <folder> --onenote-metadata
```

Each file then starts with:

```yaml
---
title: "..."            # only when it differs from the file name
created: 2025-01-01T00:00:00Z
modified: 2026-01-02T03:04:05Z
source: onenote
onenote_id: "..."
onenote_url: "..."
---
```

To make that the default, create `~/.config/onenote-export/config.json` (it
also holds your client ID) and add `"onenote_metadata": true`. A command line
flag always wins over the file.

Either way, each file's modified time is set to the page's last-modified time.
The page-to-file mapping is kept in `.onenote-export/manifest.json` at the top
of the export folder, so keep that folder if you plan to keep syncing.
Changing the setting re-exports the affected pages on the next run. Folders
exported before this option existed are treated as having metadata, so the first
run with the new default rewrites them without it.

## What gets converted

- Headings, bold, italic, strikethrough, lists, links, simple tables.
- To-do tags become `- [ ]` and `- [x]`.
- Images and attached files are downloaded to `_assets/` and linked with
  relative paths.
- Links between OneNote pages become relative Markdown links when the target
  page is part of the export.
- Tables with merged cells or nested tables are kept as HTML.

## Known limits

- Ink, handwriting, and exact page layout are not preserved.
- Password-protected sections cannot be read through the API; they are
  reported as errors and skipped.
- Microsoft's API can lag behind for a personal account, so a page created
  minutes ago may only show up on a later run.
- If a page is renamed, links to it inside pages that did not change keep the
  old path until those pages are next exported (`--force` fixes all of them).
- `--prune` removes Markdown files but leaves unused files in `_assets/`.
- Windows limits full paths to 260 characters by default. Very long titles
  nested several levels deep can exceed that, especially under `/mnt/c`; those
  pages are reported as errors and skipped, and shorter names fix them.
- File names keep accented letters. Characters Windows forbids are replaced
  or dropped, and pages with the same title get ` (2)`, ` (3)`.

## Tests

The test suite uses simulated OneNote data and needs no Microsoft account:

```bash
python3 -m unittest discover -s tests -v
```

## License

[MIT](LICENSE)
