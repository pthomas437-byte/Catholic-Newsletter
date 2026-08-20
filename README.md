# Catholic Newsletter (Blessed Sacrament Catholic Church)

A weekly newsletter summarizing news from the Vatican (Vatican News, Holy See
Press Office), the wider Catholic press (Catholic News Agency, EWTN, USCCB),
and North Carolina Catholic news (Dioceses of Raleigh & Charlotte), formatted
as a Navy (`#000080`) & Gold (`#FFD700`) `.docx` file. The banner reads:
"Catholic Newsletter" / "Blessed Sacrament Catholic Church" / the date the
issue was produced.

## How it works

1. Every Monday at 6:00 AM (America/New_York), a scheduled cloud agent
   searches the sources above, writes a summary of the week's news, and saves
   it as `content-<date>.json` in this repo.
2. It then runs `generate_newsletter.py` on that JSON to produce the formatted
   `.docx` file in the `issues/` folder, and commits/pushes both files.
3. Every Monday at 7:00 AM, a scheduled task on this PC runs `git pull` in
   this folder and copies the newest `.docx` into
   `C:\Users\pthom\iCloudDrive\Catholic Newsletter` — no manual steps needed.

### Monthly Spanish edition

Starting with the **2026-08-17** issue, the routine also produces a Latin
American Spanish (`es`) translation of that week's content once a month — on
the issue that falls on the first Monday on or after the 17th of the month
(so roughly: Aug 17, Sept 21, Oct 19, Nov 16, ...). All other weeks remain
English-only.

The Spanish edition is a separate content file with a top-level
`"language": "es-419"` field, e.g. `content-<date>-es.json`, translating the
title, headings, intro, article titles/summaries, and footer note into
Spanish (source URLs are left pointing at the original-language article).
Running `generate_newsletter.py` on it produces
`Catholic_Intelligence_Newsletter_<date>_ES.docx` — the `language` field
adds the `_ES` suffix automatically and doesn't affect the English filename.

## Folder structure

- `generate_newsletter.py` — builds the Navy & Gold `.docx` from a JSON content file.
- `content.example.json` — example of the JSON shape the generator expects (English).
- `issues/` — each week's generated `.docx` newsletter(s), plus the `content-<date>.json`
  (and, once a month, `content-<date>-es.json`) each was built from.

## Running it yourself (optional)

You don't need to do this — it happens automatically — but if you ever want
to generate an issue by hand:

1. Install Python from python.org (not the Microsoft Store version).
2. Open a terminal in this folder and run:
   ```
   pip install -r requirements.txt
   python generate_newsletter.py content.example.json
   ```
3. The `.docx` will appear in `issues/`.

## Changing the schedule or sources

The weekly search-and-build routine is managed at
https://claude.ai/code/routines — open it there to change the schedule,
sources, or wording. The local `git pull` task is a Windows Scheduled Task
named `CatholicNewsletterPull` — open Task Scheduler on this PC to adjust it.
