# DocRevise — AI Document Revision Agent

A Windows desktop tool (browser UI + local Flask server) that applies engineering
change orders to a set of documents. You give it a plain-English change request
(e.g. "replace part ABC with DEF, new pressure 60 bar"); it finds every `.docx`/`.pdf`
that mentions the old part, rewrites the wording, swaps the part image, and saves a
new revision — **without touching the originals**.

For live demos it also shows each affected document in a **Live Document View**
panel in the browser UI and animates the text + image changes on screen as they
happen — so an audience can watch the edit instead of just reading a log.
(An earlier version tried to drive an actual on-screen Word window via COM
automation, but Windows blocks background processes from stealing foreground
focus, so the window only ever showed up in the taskbar. The in-app live view
sidesteps that entirely — it's guaranteed visible since it's the same browser
tab already on screen.)

---

## What you need

| Requirement | Notes |
|---|---|
| **Windows 10/11** | The app is built for Windows, though the live view itself is just a browser panel. |
| **Python 3.9+** | Must be on PATH (`python --version` works). |
| **Internet** | The agent calls Google's Gemini API. |
| **A Gemini API key** | Get one at https://aistudio.google.com/apikey . Asked for once on first run (see below). |

Python dependencies (installed automatically by `run.bat`): `flask`, `python-docx`,
`PyMuPDF`, `requests`.

---

## Setup & run

1. Clone this repo.
2. Double-click **`run.bat`** (or run it from a terminal in the repo folder).

`run.bat` bootstraps everything:
- checks Python is installed,
- installs the Python dependencies on first run,
- asks for your **Gemini API key** the first time and saves it to a local `.env`
  file (you won't be asked again on that computer),
- starts the server and opens the app in your browser (`http://localhost:7479`).

> The API key is **not** in the repo — this is a public repo, so `.env` is gitignored.
> If you'd rather not paste it interactively, create a file named `.env` in the repo
> root containing one line:
> ```
> GEMINI_API_KEY=your_key_here
> ```

---

## How to use it

In the browser UI, fill the three fields and click **Run Agent**:

| Field | Use this |
|---|---|
| **Documents Folder** | `current_database`  (the `.docx` files to update) |
| **Change Order File** | `prompt/prompt_doc1.docx`  (the change request) |
| **Image Gallery** | `images_database`  (replacement part photos) |

`prompt/prompt_doc1.docx` asks to replace part **ABC → DEF** and update the pressure
to **60 bar** — that affects **DOC1**. Each sample prompt targets one document:
- `prompt_doc1.docx` / `prompt_doc3.docx` → ABC → DEF (affects DOC1)
- `prompt_doc2.docx` → XYZ → DEF (affects DOC3)

When you run it, for each affected document the app shows it in the **Live Document
View** panel, animates the part name/spec swap (struck-through old text, highlighted
new text) and the part photo swap right on screen, then saves a new revision next to
the original (`DOC1_RevA.docx` → `DOC1_RevB.docx`). The live log narrates every step
in plain language alongside it.

---

## Project layout

| Path | What it is |
|---|---|
| `run.bat` | One-click setup + launch (Windows). |
| `server.py` | Flask server; serves the UI and runs the pipeline (streams the log). |
| `agent.py` | The pipeline: parse change order → find docs → rewrite text → swap images → live document view. Model: `gemini-3.1-flash-lite`. |
| `ui/` | Browser front-end (`index.html`, `style.css`, `app.js`). |
| `current_database/` | The documents to update (`DOC1–3_RevA.docx`, each with an embedded part image). |
| `images_database/` | The replacement-image gallery (`Valve-*.jpg`, in `ABC/DEF/XYZ` subfolders — matching is by filename, folders are cosmetic). |
| `prompt/` | Sample change-order files. |
| `requirements.txt` | Python dependencies. |
| `.env` | Your Gemini API key (created on first run; not committed). |

---

## Notes for whoever runs this

- **Originals are never modified.** Output is always a new incremented revision saved
  in the same folder.
- **Live view is best-effort.** If it hits a problem building the preview for a
  document, the run still completes and saves the files — it just skips the
  on-screen animation for that one.
- **Close the target document in Word before running.** A document open in Word is
  locked and can't be overwritten by the save step.
- **Set `DOCREVISE_DEBUG=1`** (environment variable) to show the raw technical log lines;
  by default the log is plain-English narration for a live audience.
