# 🎓 Assignment Tutor

A Streamlit app that teaches the *reasoning* behind an assignment. A student pastes or uploads a
question (text, PDF, Word file or photo), picks a mode and a level, and gets a tutor-style
explanation from Google Gemini that they can download.

## Features

- **Four modes:** step-by-step solution, essay outline (plan only, never the essay), check my
  answer (paste or upload your own work), concept explanation.
- **Level matching:** primary school to postgraduate.
- **Input:** pasted text, PDF, DOCX, or a photo (PNG, JPG, WEBP). Phone photos are auto-rotated
  and resized. Scanned PDFs work because Gemini reads them directly.
- **Download:** Markdown (.md) and Word (.docx).
- **Error handling:** empty question, missing answer in "check" mode, oversized or corrupt
  files, missing API key, invalid key, quota/rate limits, blocked or empty AI responses,
  truncated answers, network errors, and a short cooldown against double clicks.

## Files

| File | Purpose |
|---|---|
| `app.py` | The whole app |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Get a free API key at <https://aistudio.google.com/app/apikey>, then create
`.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "your-key-here"
# GEMINI_MODEL = "gemini-3.5-flash"   # optional: use a different Gemini model
```

```bash
streamlit run app.py
```

(Alternatively set the `GEMINI_API_KEY` environment variable.)

## Push to GitHub (web UI only)

1. Sign in at <https://github.com> → click **+** (top right) → **New repository**.
2. Name it (for example `assignment-tutor`), choose Public or Private, leave "Add a README"
   **unticked**, click **Create repository**.
3. On the empty repo page click **uploading an existing file**. Drag in `app.py`,
   `requirements.txt` and `README.md`. Keep the commit option on "Commit directly to the
   main branch" and click **Commit changes**.
4. Click **Add file → Create new file**, name it `.gitignore`, paste the lines below, and
   commit. This keeps secrets out of the repo.
   ```
   .streamlit/secrets.toml
   .env
   __pycache__/
   .venv/
   ```

**Never upload your API key or `secrets.toml` to GitHub.** Keys in a repo get found and abused.

## Deploy on Streamlit Community Cloud

1. Go to <https://share.streamlit.io> and sign in with GitHub. Authorize access to your repo
   if asked (needed for private repos too).
2. Click **Create app** (or **New app**) → deploy from an existing repo.
3. Pick your repository, branch `main`, and main file path `app.py`.
4. Open **Advanced settings**. Choose Python 3.11 or 3.12, and in **Secrets** paste:
   ```toml
   GEMINI_API_KEY = "your-key-here"
   ```
5. Click **Deploy**. The first build takes a few minutes. You get a public
   `*.streamlit.app` URL.
6. To change the key later: open the app → **Manage app** → **Settings** → **Secrets**.
   To update the code: edit a file on GitHub and commit; the app redeploys automatically.

(Button names in Streamlit's dashboard change from time to time; the flow stays the same.)

## Configuration

| Secret | Required | Meaning |
|---|---|---|
| `GEMINI_API_KEY` (or `GOOGLE_API_KEY`) | yes | Your Gemini API key |
| `GEMINI_MODEL` | no | Model id; default is `gemini-3.5-flash`. Change it here when Google releases or retires models. |

Limits live at the top of `app.py`: `MAX_FILE_MB` (15), `MAX_TEXT_CHARS` (20,000),
`COOLDOWN_SECONDS` (5), `MAX_OUTPUT_TOKENS`.

## Troubleshooting

- **"no Gemini API key was found"**: the secret is missing or misspelled. Names are case-sensitive.
- **"model was not found"**: set `GEMINI_MODEL` to a model your key can use.
- **"busy or the free quota has been used up"**: wait a minute, or enable billing in Google AI Studio.
- **"declined to process" / "blocked"**: Gemini's safety filter fired, sometimes wrongly on
  homework. Reword the question.
- **Maths looks odd in the Word file**: Word can't render LaTeX, so formulas appear as plain
  `$...$` text there. The on-screen view and the Markdown file render them properly in Streamlit
  or any Markdown viewer that supports maths.
- **Scanned Word files**: a DOCX that only contains pictures has no text to read. Upload it as
  a PDF or photo instead.

## Notes for the app owner

- **Cost and abuse:** a public app spends *your* quota. Keep the free tier, set a budget in
  Google Cloud, or restrict who can open the app (Streamlit Community Cloud lets you limit
  viewers by email in the app's sharing settings).
- **Privacy:** questions and uploads are sent to Google's Gemini API. Tell students not to
  upload personal information, and check Google's current terms for how data is handled on the
  free versus paid tier. Nothing is stored by the app itself; results live only in the browser
  session.
- **Academic integrity:** the prompts are written to teach and to plan, not to produce
  submit-ready work (the essay mode never writes the essay).
- **Optional:** to reject huge uploads earlier, add `.streamlit/config.toml` with
  `[server]` and `maxUploadSize = 15`.
