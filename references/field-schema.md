# candidate.json schema — Ford/Magnit CV Template

Feed this JSON into `scripts/build_cv.py --data candidate.json`. Every field
is optional; anything omitted falls back to `"NA"` (Ford-history / cert /
score fields) or `"TBC"` (scheduling fields), matching the convention
Rajkumar asked for in Sep 2026.

## Extraction rule

Use ONLY information present in the resume. Do not invent, infer, or
estimate anything not stated. If a field is not present in the resume, use
`"NA"` (or `"TBC"` for the four scheduling fields below) rather than
guessing.

## Top-level fields

| Key | Notes |
|---|---|
| `first_name` | |
| `last_name` | |
| `notice_period` | Keep SHORT — see "Line-wrap warning" below. Default `"TBC"`. |
| `internal_external` | `"Internal"` / `"External"` / default `"TBC"` |
| `interview_availability` | Keep SHORT. Default `"TBC"` |
| `start_availability` | Keep SHORT. Default `"TBC"` |
| `worked_ford_direct` | `"Yes"` / `"No"`. Default `"No"` if not stated. |
| `worked_ford_agency` | `"Yes"` / `"No"`. Default `"No"` if not stated. |
| `cdsid` | Only relevant if either of the above is Yes. Default `"NA"` |
| `supervision_duration` | Default `"NA"` |
| `exit_reason` | Default `"NA"` |
| `overall_it_experience` | Only fill if the resume states a total, e.g. "6 Years". Default `"NA"` |
| `experience_core_skill` | e.g. "5 Years". Look beyond a literal "core skill experience" label — check the summary/profile section and the top_skills-relevant employment entries for a stated total against the candidate's primary/most-recent skill (e.g. "5+ years in Data Engineering", "3 yrs on AWS"). Only use `"NA"` if the resume genuinely states nothing usable, not just because no field is explicitly labeled this. |
| `other_qualifications` | Default `"NA"` |
| `hackerrank_score` | Default `"NA"` |
| `certifications` | One string, multiple certs separated by `;`, e.g. `"DP-700: ...; DP-200: ...; MCSE ..."`. The script numbers each one onto its own line automatically — don't pre-number them yourself. Default `"NA"` |
| `top_skills` | List of exactly 3 short phrases (not sentences), most prominent skills/domains from the resume. |
| `education_bachelor` | One line: degree – college (affiliation) – completion date. |
| `education_master` | One line: degree(s) (specialization) – university name. If candidate has 2 postgrad degrees (e.g. M.Sc + M.Tech), combine with `;`. |
| `education_master_duration` | e.g. "NA – 05.2022 (M.Tech); NA – 06.2020 (M.Sc)" — resume rarely states a start date, so leading side is usually "NA". |
| `education_bachelor_detail` | Repeats bachelor's degree, college, and university affiliation (the template has a separate line for this). |
| `employment` | List of jobs, **most recent first**, see below. Not capped at 3 — the script rebuilds the whole section for however many entries you give it. |
| `resume` | Optional. If present, the resume is appended after the template as **editable text** (not an image). See "Resume" below for the schema and what must be excluded. |

### `employment[]` entry

```json
{
  "duration": "June 2022 \u2013 Present",
  "company": "Hewlett Packard (HP Inc), Bengaluru, India",
  "title": "Embedded System Engineer & AI Specialist",
  "technology": "Agentic AI, MCP server, Embedded Linux, Yocto Project",
  "role": ["Implemented LLM automation in CI/CD ...", "Led migration of ...", "..."]
}
```

`role` should normally be a **list of short bullet points** (3-6 lines,
condensed from the resume's own bullets — don't rewrite the meaning), not
one long paragraph — recruiters scan, they don't read paragraphs
(Rajkumar's Sep-2026 feedback). A plain string is still accepted for a
one-line role, but prefer the list form whenever the resume gives more
than a line's worth of detail.

**Company-name consistency check:** before filling `employment[]`, verify
every `company` value against the source resume's own wording. Because
client background-verification teams check this, a company name that
doesn't match the resume verbatim (wrong employer, merged/renamed entries,
etc.) is a serious defect — re-read the resume's employment section and
fix any mismatch before running the script, don't just trust an
in-progress candidate.json.

### Resume (appended as editable text — contact details, photo, and logos excluded)

Standing instruction from Rajkumar (Sep 2026): the appended resume must be
**editable Word text**, not a picture of the original file, and must
**never include**:
- the candidate's contact details — email, mobile number, LinkedIn,
  GitHub, or any other personal link (TEKsystems/Magnit handle candidate
  contact, so this doesn't go to the client)
- the candidate's photo
- any organisation/company logo

The candidate's **name** is fine to keep — it's not a contact method.
Everything else (roles, dates, bullets, education, skills, projects,
certificates) should be reproduced faithfully from the resume, condensed
lightly if needed but not rewritten in meaning — same rule as the
`employment[]` entries above.

```json
{
  "name": "AMESH A M",
  "summary": ["Optional — one string or a list of bullet lines, if the resume has a career-summary/profile section worth keeping."],
  "work_experience": [
    {"title": "...", "company": "...", "duration": "...", "bullets": ["...", "..."]}
  ],
  "education": ["MBA/PGDM \u2014 P E S Institute of Technology, Bangalore \u2014 2019", "..."],
  "skills": ["Roadmap", "Backlog Refinement", "Scrum", "..."],
  "projects": [{"title": "...", "description": "..."}],
  "certificates": ["POPM", "Google Data Analytics Professional Certificate", "..."],
  "languages": ["English", "Kannada", "Hindi"]
}
```

Do not add an `"email"`, `"phone"`, `"linkedin"`, `"github"`, `"contact"`,
or `"photo"` key to this block — `build_resume_section_xml()` in
`scripts/build_cv.py` doesn't render one even if present, but the cleanest
approach is to just never extract those fields from the resume into this
JSON in the first place.


## Line-wrap warning

`notice_period`, `internal_external`, `interview_availability`, and
`start_availability` sit in a text box that reserves right-side space for
the candidate photo. Long values (e.g. the literal phrase "To be
confirmed") wrap onto a second line and look broken. Use short values:
`"TBC"`, `"Immediate"`, `"7 Days"`, `"15 Days"`, `"2 Weeks"`. If a value
must be longer, check the rendered PDF and shorten it if it wraps.

## Photo (CV Template's own field only)

As of the Sep-2026 template revision, the CV Template's photo box (page 1)
ships **empty** — Ford/Magnit removed the old sample placeholder photo.
Pass `--photo /path/to/photo.jpg` to insert a real candidate photo into
that box. It gets center-cropped to the box's aspect ratio and inserted —
see `insert_photo()` in `scripts/build_cv.py` for how this works (it
places a new picture into the template's empty photo-box shape, since
there's no existing image to overwrite anymore).

Per Rajkumar's Sep-2026 instruction, don't source this photo from the
candidate's resume automatically — only pass `--photo` when he's
explicitly supplied one for this field. If `--photo` is omitted, the box
stays empty, which is now the correct default (not a sign something's
missing) — no need to flag it to Rajkumar unless he's separately supplied
a photo you forgot to pass in. (The resume section appended below the
template never has a photo, regardless.)
