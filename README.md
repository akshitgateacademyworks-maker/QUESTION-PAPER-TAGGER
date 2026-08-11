# MCQ Docx Tagger

A small Streamlit app that converts a plain (untagged) MCQ Word document
into the `@que / @body / @options / @ans / @ep` **tagged** format used for
question-bank uploads:

```
$total: <N>
@que
  @type
    Option
  !type
  @body
    <question text>
  !body
  @options
    a) option 1
    b) option 2
    c) option 3
    d) option 4
  !options
  @ans
    C
  !ans
  @ep
    <explanation>
  !ep
!que
```

Tag lines render in red Arial, content in black Times New Roman, and the
correct-answer letter is bold — matching the reference tagging template.

## How it works

- `tagger.py` — parses the uploaded plain docx (regex-based: numbered
  questions, lettered options, an `Answer:`/`Ans:` line, an optional
  `Explanation:`/`Solution:` line) and writes a new `.docx` replicating the
  tag structure, including a real Word lettered list for options that
  restarts at `a)` for every question.
- `app.py` — the Streamlit front end: upload → preview parsed questions →
  download the tagged `.docx`.
- `assets/template_shell.docx` — a stripped-down copy of the reference
  tagged document (styles/fonts only, no question content) that the
  generator reuses so fonts/styles match exactly. If this file is ever
  missing, `tagger.py` falls back to building a minimal docx package from
  scratch.

## Input format the parser expects

```
1. Which type of sewer is specifically designed to handle a discharge
   25 times greater than the sanitary discharge?
a) Separate sewer
b) Outfall sewer
c) Combined sewer
d) Lateral sewer
Answer: C
Explanation: Combined sewers handle both sanitary sewage and storm
drainage, which is much larger.
```

Numbering/lettering punctuation is flexible (`1.`, `Q1)`, `a)`, `(a)`, `Ans:`,
`Answer -`, `Explanation:`, `Solution:` are all recognised). Blank lines
between questions are fine. Text between the question number and the first
option line becomes a multi-paragraph question body (useful for
assertion/reason style questions). Questions with no options detected are
tagged `Numerical` instead of `Option` and skip the `@options` block.

A ready-made example is in `sample_files/sample_untagged.docx`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

### Push to GitHub

```bash
git init
git add .
git commit -m "MCQ docx tagger"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

### Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with
   GitHub.
2. Click **New app**, pick this repo, branch `main`, and set the main file
   to `app.py`.
3. Deploy — no secrets or extra configuration needed.

## Notes / limitations

- The parser is regex-based and tuned to the format above. If your source
  documents use a very different layout (e.g. no explicit `Answer:` label,
  tables, images inside questions), tell me the exact format and I'll adjust
  `parse_plain_docx`.
- Only the `Option` question type is generated with a lettered list;
  questions without recognisable options are tagged `Numerical` with no
  `@options` block, per the pattern seen in typical tagging schemes — adjust
  `Question.qtype` logic in `tagger.py` if your system uses other type names.
