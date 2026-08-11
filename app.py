import os
import tempfile

import streamlit as st

from tagger import convert, parse_plain_docx

st.set_page_config(page_title="MCQ Tagger", page_icon="🏷️", layout="centered")

st.title("🏷️ MCQ Docx Tagger")
st.write(
    "Upload a plain (untagged) MCQ Word document and get back a copy in the "
    "`@que / @body / @options / @ans / @ep` tagging format used for "
    "question-bank uploads."
)

with st.expander("Expected input format"):
    st.markdown(
        """
Each question should look roughly like this (numbers, letters and labels
are flexible — `1.`, `Q1)`, `a)`, `(a)`, `Ans:`, `Answer -`, `Explanation:`,
`Solution:` are all recognised):

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

2. Next question...
```

Blank lines between questions are fine. Multi-line question bodies
(e.g. assertion/reason style) are supported — everything between the
question number and the first option line becomes the question body.
        """
    )

uploaded = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded is not None:
    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input.docx")
        out_path = os.path.join(tmp, "tagged_output.docx")
        with open(in_path, "wb") as f:
            f.write(uploaded.getvalue())

        try:
            questions = parse_plain_docx(in_path)
        except Exception as e:
            st.error(f"Couldn't read that file: {e}")
            questions = []

        if not questions:
            st.warning(
                "No questions were detected. Check that questions start with "
                "a number (e.g. `1.`) and options start with a letter "
                "(e.g. `a)`) — see the expected format above."
            )
        else:
            st.success(f"Detected {len(questions)} question(s).")

            with st.expander("Preview parsed questions"):
                for i, q in enumerate(questions, start=1):
                    st.markdown(f"**Q{i}.** {' '.join(q.body)}")
                    for letter, opt in zip("abcdefgh", q.options):
                        marker = "✅" if letter.upper() == (q.answer_letter or "") else "•"
                        st.write(f"{marker} {letter}) {opt}")
                    if q.explanation:
                        st.caption("Explanation: " + " ".join(q.explanation))
                    st.divider()

            convert(in_path, out_path)
            with open(out_path, "rb") as f:
                st.download_button(
                    "⬇️ Download tagged .docx",
                    data=f.read(),
                    file_name=f"tagged_{uploaded.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
else:
    st.info("Waiting for a .docx upload.")
    sample_path = os.path.join(os.path.dirname(__file__), "sample_files", "sample_untagged.docx")
    if os.path.exists(sample_path):
        with open(sample_path, "rb") as f:
            st.download_button(
                "Download a sample untagged .docx to try",
                data=f.read(),
                file_name="sample_untagged.docx",
            )
