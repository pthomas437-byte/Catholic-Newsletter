# Catholic Intelligence Newsletter

A weekly newsletter summarizing news from the Vatican (Vatican News, Holy See
Press Office) and the wider Catholic press (Catholic News Agency, EWTN,
USCCB), formatted as a Navy & Gold `.docx` file.

## How it works

1. Every Monday at 6:00 AM (America/New_York), a scheduled cloud agent
   searches the sources above, writes a summary of the week's news, and saves
   it as `content-<date>.json` in this repo.
2. It then runs `generate_newsletter.py` on that JSON to produce the formatted
   `.docx` file in the `issues/` folder, and commits/pushes both files.
3. Every Monday at 7:00 AM, a scheduled task on this PC runs `git pull` in
   this folder, so the new issue shows up here automatically — no manual
   steps needed.

## Folder structure

- `generate_newsletter.py` — builds the Navy & Gold `.docx` from a JSON content file.
- `content.example.json` — example of the JSON shape the generator expects.
- `issues/` — each week's generated `.docx` newsletter, plus the `content-<date>.json` it was built from.

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
