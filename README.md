# Ford / Magnit CV Generator

This Streamlit application wraps the supplied Ford/Magnit CV skill package.

## Run locally

```bash
python -m venv .venv
.venv\\Scripts\\activate   # Windows
pip install -r requirements.txt
streamlit run app.py
```

On macOS/Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Workflow

1. Upload PDF, DOCX or TXT resume.
2. Review extracted text.
3. Populate/review `candidate.json` using the supplied schema.
4. Optionally upload a separate candidate photo.
5. Generate the editable Word CV.

## Important

The bundled `build_cv.py` relies on the exact paragraph structure of the supplied `Candidate_Template.docx`. If the template changes, anchors may need to be updated.

The app intentionally does not automatically extract a candidate photo from the resume.
