"""
XML Job Feed Generator
======================

A local Streamlit dashboard that turns job form input + a plain-text job
description into a valid XML feed:

    <source>
        <job>
            <job_title>...</job_title>
            ...
            <job_description><![CDATA[ <p>...clean HTML...</p> ]]></job_description>
        </job>
    </source>

Run with:
    pip install streamlit lxml pandas openpyxl
    streamlit run xmlapp.py
"""

import os
import re
import json
import textwrap
import urllib.request
import html as html_lib

import streamlit as st
import pandas as pd
from lxml import etree

# --------------------------------------------------------------------------- #
# CONFIGURATION
# Each field becomes a tag inside <job>. `name` is the XML tag, `label` is the
# form label, `required` enforces input before a job can be added.
# --------------------------------------------------------------------------- #
FIELDS = [
    {"name": "reference_number", "label": "Reference Number (Job ID)", "required": True},
    {"name": "job_title",        "label": "Job Title",                 "required": True},
    {"name": "company",          "label": "Company",                   "required": True},
    {"name": "job_url",          "label": "Apply / Job URL",           "required": True},
    {"name": "job_state",        "label": "State / Region",            "required": True},
    {"name": "job_country",      "label": "Country",                   "required": True},
    # ---- Optional fields -------------------------------------------------- #
    {"name": "salary",           "label": "Salary",                    "required": False},
    {"name": "city",             "label": "City",                      "required": False},
    {"name": "job_category",     "label": "Job Category",              "required": False},
    {"name": "zipcode",          "label": "Zipcode",                   "required": False},
]
DESCRIPTION_FIELD = "job_description"  # the (optionally) CDATA-wrapped HTML tag

FIELD_KEYS = [f"in_{f['name']}" for f in FIELDS] + ["in_description"]

# --------------------------------------------------------------------------- #
# HARD-CODED SIDEBAR IMAGES (permanent, built-in — no uploads).
#
# Paste the *raw* base64 of each image between the quotes below (no "data:..."
# prefix, no quotes inside). Quick way to generate it from a file:
#     python -c "import base64,sys;print(base64.b64encode(open(sys.argv[1],'rb').read()).decode())" me.png
# Copy the whole printed string and paste it as PROFILE_PIC_BASE64.
#
# IMAGE_MIME tells the browser the format. Use "image/png", "image/jpeg",
# or "image/gif" to match each image (a GIF will animate on hover).
#
# Until you replace the PASTE_..._HERE placeholders, the app falls back to the
# placeholder URLs below so nothing renders as a broken image.
# --------------------------------------------------------------------------- #
PROFILE_PIC_BASE64 = "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"
HOVER_CAT_BASE64 = "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"

PROFILE_MIME = "image/jpeg"  # baked-in profile picture (JPEG)
HOVER_MIME = "image/jpeg"    # baked-in hover cat (JPEG)

# Fallbacks used only while the base64 placeholders above are still unset.
PROFILE_IMG_FALLBACK = "https://placehold.co/120x120/4F46E5/ffffff?text=Me"
SAD_CAT_FALLBACK = "https://cataas.com/cat/says/Sad%20Cat?fontSize=28&fontColor=white"


# --------------------------------------------------------------------------- #
# PURE LOGIC  (no Streamlit calls -> independently testable)
# --------------------------------------------------------------------------- #
_NUM_RE = re.compile(r"^\s*\d+[.)]\s+(.*)$")
_BUL_RE = re.compile(r"^\s*[-*•▪◦·]\s+(.*)$")
_HDR_RE = re.compile(r"^\s*(#{1,6})\s+(.*)$")   # markdown headers:  #, ##, ### ...
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")          # inline bold:  **text**


def _fmt(text: str) -> str:
    """Escape one line of user text, then turn **bold** into <strong>.

    Escaping runs first so any literal < > & in the job text stay valid XML/HTML.
    The ** markers contain no special characters, so they survive escaping and
    are converted to <strong> afterwards.
    """
    esc = html_lib.escape(text)
    return _BOLD_RE.sub(r"<strong>\1</strong>", esc)


def _classify(line: str):
    """Return ('ul'|'ol'|'p', content) for a single (non-header) line."""
    m = _BUL_RE.match(line)
    if m:
        return "ul", m.group(1).strip()
    m = _NUM_RE.match(line)
    if m:
        return "ol", m.group(1).strip()
    return "p", line.strip()


def text_to_html(text: str) -> str:
    """Convert plain text into clean HTML.

    Supported markup (type it in the plain-text box):
      # / ## / ### line   ->  <h1>..<h6> header
      **bold**            ->  <strong>bold</strong>  (anywhere on a line)
      - or * or • line    ->  <ul><li>..</li></ul> bullet list
      1. or 2) line       ->  <ol><li>..</li></ol> numbered list
      blank line          ->  new block / paragraph
      anything else       ->  <p> (consecutive lines joined with <br/>)
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return ""

    html_parts = []

    for block in re.split(r"\n\s*\n", text):
        lines = [ln for ln in block.split("\n") if ln.strip()]
        if not lines:
            continue

        run, run_type = [], None

        def flush():
            if not run:
                return
            items = [_fmt(x) for x in run]
            if run_type == "ul":
                html_parts.append("<ul>" + "".join(f"<li>{x}</li>" for x in items) + "</ul>")
            elif run_type == "ol":
                html_parts.append("<ol>" + "".join(f"<li>{x}</li>" for x in items) + "</ol>")
            else:
                html_parts.append("<p>" + "<br/>".join(items) + "</p>")

        for ln in lines:
            hdr = _HDR_RE.match(ln)
            if hdr:
                flush()
                run, run_type = [], None
                level = len(hdr.group(1))
                html_parts.append(f"<h{level}>{_fmt(hdr.group(2).strip())}</h{level}>")
                continue

            kind, content = _classify(ln)
            if kind != run_type:
                flush()
                run, run_type = [], kind
            run.append(content)
        flush()

    return "\n".join(html_parts)


def _cdata(value: str):
    """Safe CDATA wrapper. ']]>' is illegal inside CDATA, so we neutralise it."""
    return etree.CDATA(value.replace("]]>", "]] >"))


def build_xml(jobs, wrap_desc_cdata=True, wrap_url_cdata=False,
              publisher=None, publisher_url=None) -> bytes:
    """Build the full feed and return pretty-printed UTF-8 XML bytes.

    wrap_desc_cdata : wrap the <job_description> tag in CDATA (raw HTML preserved).
                      If False, the HTML is auto-escaped instead (still valid XML).
    wrap_url_cdata  : wrap only the <job_url> tag in CDATA.
    """
    root = etree.Element("source")

    if publisher:
        etree.SubElement(root, "publisher").text = publisher
    if publisher_url:
        etree.SubElement(root, "publisherurl").text = publisher_url
    # (lastBuildDate node intentionally removed.)

    for job in jobs:
        job_el = etree.SubElement(root, "job")

        for f in FIELDS:
            value = (job["fields"].get(f["name"], "") or "").strip()
            el = etree.SubElement(job_el, f["name"])
            if not value:
                continue  # leave as empty <tag/>

            # Only job_url is CDATA-wrapped, and only if its checkbox is on.
            # lxml auto-escapes & < > in plain text, so plain text is valid XML.
            use_cdata = (f["name"] == "job_url" and wrap_url_cdata)
            el.text = _cdata(value) if use_cdata else value

        desc_el = etree.SubElement(job_el, DESCRIPTION_FIELD)
        if wrap_desc_cdata:
            desc_el.text = _cdata(job["description_html"])
        else:
            desc_el.text = job["description_html"]  # lxml escapes the HTML

    return etree.tostring(
        root, pretty_print=True, xml_declaration=True, encoding="UTF-8"
    )


# --------------------------------------------------------------------------- #
# BULK CSV / EXCEL  (pure logic -> independently testable)
# --------------------------------------------------------------------------- #
ALL_COLUMNS = [f["name"] for f in FIELDS] + [DESCRIPTION_FIELD]


def template_csv() -> bytes:
    """A blank CSV with exactly the columns this app expects."""
    return (",".join(ALL_COLUMNS) + "\n").encode("utf-8")


def read_table(file_obj, filename: str) -> pd.DataFrame:
    """Read a CSV or Excel upload into a string DataFrame (no numeric coercion)."""
    name = (filename or "").lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj, dtype=str)
    else:
        df = pd.read_csv(file_obj, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _cell(row, col):
    """Safe single-cell read: '' for missing column or NaN."""
    if not col or col not in row.index:
        return ""
    val = row[col]
    if pd.isna(val):
        return ""
    return str(val).strip()


def build_jobs_from_df(df: pd.DataFrame, mapping: dict, desc_is_html: bool):
    """Turn a DataFrame into job dicts using a {tag_name: column_name} mapping."""
    jobs, skipped = [], []
    for idx, row in df.iterrows():
        fields, missing = {}, []
        for f in FIELDS:
            val = _cell(row, mapping.get(f["name"]))
            fields[f["name"]] = val
            if f["required"] and not val:
                missing.append(f["name"])

        raw_desc = _cell(row, mapping.get(DESCRIPTION_FIELD))
        if not raw_desc:
            missing.append(DESCRIPTION_FIELD)

        if missing:
            skipped.append((idx + 2, missing))  # +2: 1-based + header row
            continue

        desc_html = raw_desc if desc_is_html else text_to_html(raw_desc)
        jobs.append({"fields": fields, "description_html": desc_html})

    return jobs, skipped


def guess_column(target: str, columns) -> str:
    """Best-guess CSV column for a given tag name (case/underscore-insensitive)."""
    norm = lambda s: re.sub(r"[^a-z0-9]", "", str(s).lower())
    t = norm(target)
    for c in columns:
        if norm(c) == t:
            return c
    for c in columns:
        if t in norm(c) or norm(c) in t:
            return c
    return ""


# --------------------------------------------------------------------------- #
# STREAMLIT UI HELPERS
# --------------------------------------------------------------------------- #
def _init_state():
    for key in FIELD_KEYS:
        st.session_state.setdefault(key, "")
    st.session_state.setdefault("jobs", [])
    st.session_state.setdefault("xml_bytes", None)
    st.session_state.setdefault("bulk_xml_bytes", None)
    st.session_state.setdefault("flash", None)


def _read_form():
    """Return (job_dict, errors) from the current widget values."""
    fields = {f["name"]: st.session_state[f"in_{f['name']}"].strip() for f in FIELDS}
    raw_desc = st.session_state["in_description"]

    errors = [f["label"] for f in FIELDS if f["required"] and not fields[f["name"]]]
    if not raw_desc.strip():
        errors.append("Job Description")

    job = {"fields": fields, "description_html": text_to_html(raw_desc)}
    return job, errors


def _clear_form():
    for key in FIELD_KEYS:
        st.session_state[key] = ""


def _add_job():
    job, errors = _read_form()
    if errors:
        st.session_state.flash = ("error", "Missing required: " + ", ".join(errors))
        return
    st.session_state.jobs.append(job)
    st.session_state.flash = ("success", f"Added '{job['fields']['job_title']}' to the feed.")
    _clear_form()


def _clear_feed():
    st.session_state.jobs = []
    st.session_state.xml_bytes = None


def _hardcoded_src(b64: str, mime: str, fallback: str) -> str:
    """Build an <img src> from a hard-coded base64 string.

    If the base64 hasn't been pasted yet (still the PASTE_..._HERE placeholder),
    return the fallback URL so the image isn't broken.
    """
    if not b64 or b64.startswith("PASTE_"):
        return fallback
    return f"data:{mime};base64,{b64}"


def _sidebar_footer():
    """Bottom-of-sidebar profile card with a hover-to-reveal sad cat.

    Both images are hard-coded base64 (see the constants at the top of the
    file). Two images are stacked; CSS `:hover` fades the profile out and the
    cat in.
    """
    profile_src = _hardcoded_src(PROFILE_PIC_BASE64, PROFILE_MIME, PROFILE_IMG_FALLBACK)
    cat_src = _hardcoded_src(HOVER_CAT_BASE64, HOVER_MIME, SAD_CAT_FALLBACK)

    st.markdown(
        textwrap.dedent(f"""
        <style>
        .profile-card {{ text-align: center; margin-top: 1.2rem; }}
        .profile-card .bubble {{
            display: inline-block;
            background: #262730;
            color: #fafafa;
            padding: 8px 12px;
            border-radius: 12px;
            font-size: 0.82rem;
            line-height: 1.3;
            margin-bottom: 12px;
            max-width: 220px;
        }}
        .profile-card .pic {{
            position: relative;
            width: 80px;
            height: 80px;
            margin: 0 auto;
            cursor: pointer;
        }}
        .profile-card .pic img {{
            position: absolute;
            top: 0; left: 0;
            width: 80px;
            height: 80px;
            border-radius: 50%;
            object-fit: cover;
            transition: opacity 0.25s ease;
        }}
        .profile-card .pic img.cat {{ opacity: 0; }}
        .profile-card .pic:hover img.me {{ opacity: 0; }}
        .profile-card .pic:hover img.cat {{ opacity: 1; }}
        .profile-card .hint {{
            font-size: 0.72rem;
            color: #9a9aa5;
            margin-top: 8px;
        }}
        </style>

        <div class="profile-card">
            <div class="bubble">"I was tired of creating them manually"</div>
            <div class="pic">
                <img class="me"  src="{profile_src}" alt="profile" />
                <img class="cat" src="{cat_src}"  alt="sad cat" />
            </div>
            <div class="hint">Amrita G.</div>
        </div>
        """),
        unsafe_allow_html=True,
    )



# --------------------------------------------------------------------------- #
# USAGE COUNTER
# Counts one visit per browser session. Stored in a free hosted counter
# (abacus) so the number survives app restarts/redeploys; falls back to a
# local file if that service can't be reached.
# --------------------------------------------------------------------------- #
COUNTER_API = "https://abacus.jasoncameron.dev"
COUNTER_NS = "joveo-xmlfeedgenerator"
COUNTER_KEY = "visits"
# Added to the raw count so the display starts at 12 (first new visit -> 12).
COUNTER_OFFSET = 11
_LOCAL_COUNTER = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".visit_count")


def _remote_counter(action: str):
    url = f"{COUNTER_API}/{action}/{COUNTER_NS}/{COUNTER_KEY}"
    try:
        with urllib.request.urlopen(url, timeout=3) as resp:
            return int(json.loads(resp.read().decode("utf-8"))["value"])
    except Exception:
        return None


def _local_counter(increment: bool) -> int:
    try:
        with open(_LOCAL_COUNTER) as fh:
            n = int(fh.read().strip() or 0)
    except Exception:
        n = 0
    if increment:
        n += 1
        try:
            with open(_LOCAL_COUNTER, "w") as fh:
                fh.write(str(n))
        except Exception:
            pass
    return n


def get_visit_count() -> int:
    """Increment once per session, then reuse the cached value on reruns."""
    if "visit_count" not in st.session_state:
        n = _remote_counter("hit")
        if n is None:
            n = _local_counter(increment=True)
        st.session_state.visit_count = n + COUNTER_OFFSET
    return st.session_state.visit_count

# --------------------------------------------------------------------------- #
# MAIN APP
# --------------------------------------------------------------------------- #
def main():
    st.set_page_config(page_title="XML Job Feed Generator", page_icon="🧩", layout="wide")
    _init_state()

    # ---- Sidebar -------------------------------------------------------- #
    with st.sidebar:
        st.header("⚙️ Feed settings")
        filename = st.text_input("Output file name", value="job_feed.xml")
        if not filename.endswith(".xml"):
            filename += ".xml"
        publisher = st.text_input("Publisher (optional)", value="")
        publisher_url = st.text_input("Publisher URL (optional)", value="")

        wrap_desc = st.checkbox(
            "Wrap Job Description with CDATA", value=True,
            help="On = <job_description> keeps raw HTML inside CDATA. Off = the HTML is escaped instead.",
        )
        wrap_url = st.checkbox(
            "Wrap Job URL with CDATA", value=False,
            help="On = wrap only the <job_url> tag in CDATA.",
        )

        st.divider()
        st.metric("Jobs staged in feed", len(st.session_state.jobs))
        st.button("🗑️ Clear feed", on_click=_clear_feed, use_container_width=True)

        st.divider()
        # Permanent, built-in profile card pinned at the very bottom.
        _sidebar_footer()

    # ---- Header --------------------------------------------------------- #
    head_l, head_r = st.columns([4, 1])
    with head_l:
        st.title("🧩 XML Job Feed Generator")
    with head_r:
        st.metric("👥 Users so far", f"{get_visit_count():,}")
    st.caption(
        "Generate a valid `<source>` / `<job>` feed. Choose per-tag CDATA "
        "wrapping in the sidebar."
    )

    if st.session_state.flash:
        level, msg = st.session_state.flash
        getattr(st, level)(msg)
        st.session_state.flash = None

    tab_manual, tab_bulk = st.tabs(["✍️ Single / manual entry", "📤 Bulk CSV / Excel import"])

    with tab_manual:
        _render_manual_tab(filename, publisher, publisher_url, wrap_desc, wrap_url)

    with tab_bulk:
        _render_bulk_tab(filename, publisher, publisher_url, wrap_desc, wrap_url)


def _render_manual_tab(filename, publisher, publisher_url, wrap_desc, wrap_url):
    # ---- Input form ----------------------------------------------------- #
    st.subheader("1 · Job details")
    cols = st.columns(3)
    for i, f in enumerate(FIELDS):
        label = f["label"] + (" *" if f["required"] else "")
        cols[i % 3].text_input(label, key=f"in_{f['name']}")

    st.subheader("2 · Job description (plain text)")
    st.caption(
        "Formatting you can type:  `# Heading` (also `##`, `###`)  ·  "
        "`**bold text**`  ·  `- ` bullet list  ·  `1.` numbered list  ·  "
        "blank line = new paragraph."
    )
    st.text_area(
        "Paste or type the description here.",
        key="in_description",
        height=260,
    )

    c1, c2, _ = st.columns([1, 1, 2])
    c1.button("➕ Add job to feed", on_click=_add_job, use_container_width=True)
    c2.button("🧹 Clear form", on_click=_clear_form, use_container_width=True)

    if st.session_state["in_description"].strip():
        with st.expander("👁️ Preview generated HTML for current description"):
            html = text_to_html(st.session_state["in_description"])
            st.code(html, language="html")
            st.markdown("**Rendered:**")
            st.markdown(html, unsafe_allow_html=True)

    # ---- Staged feed ---------------------------------------------------- #
    if st.session_state.jobs:
        st.subheader("3 · Staged jobs")
        st.dataframe(
            [
                {
                    "Job Title": j["fields"].get("job_title", ""),
                    "Company": j["fields"].get("company", ""),
                    "State": j["fields"].get("job_state", ""),
                    "Ref #": j["fields"].get("reference_number", ""),
                }
                for j in st.session_state.jobs
            ],
            use_container_width=True,
            hide_index=True,
        )

    # ---- Generate ------------------------------------------------------- #
    st.subheader("4 · Generate XML")
    st.caption(
        "Uses every staged job. If none are staged, the current form is used as a "
        "single job (it must pass required-field validation)."
    )

    if st.button("⚡ Generate XML", type="primary", key="gen_manual"):
        jobs = list(st.session_state.jobs)
        if not jobs:
            job, errors = _read_form()
            if errors:
                st.error("Nothing staged and the form is incomplete. Missing: "
                         + ", ".join(errors))
                st.session_state.xml_bytes = None
            else:
                jobs = [job]
        if jobs:
            st.session_state.xml_bytes = build_xml(
                jobs, wrap_desc_cdata=wrap_desc, wrap_url_cdata=wrap_url,
                publisher=publisher or None, publisher_url=publisher_url or None,
            )
            st.success(f"Generated feed with {len(jobs)} job(s).")

    if st.session_state.xml_bytes:
        with st.expander("📄 XML preview", expanded=True):
            st.code(st.session_state.xml_bytes.decode("utf-8"), language="xml")
        st.download_button(
            "⬇️ Download .xml",
            data=st.session_state.xml_bytes,
            file_name=filename,
            mime="application/xml",
            type="primary",
            use_container_width=True,
            key="dl_manual",
        )


def _render_bulk_tab(filename, publisher, publisher_url, wrap_desc, wrap_url):
    st.subheader("1 · Get the template (optional)")
    st.caption(
        "Your spreadsheet needs one column per field. Name them like the template "
        "and they'll auto-map. Any column names work too — map them by hand below."
    )
    st.download_button(
        "⬇️ Download CSV template",
        data=template_csv(),
        file_name="job_feed_template.csv",
        mime="text/csv",
    )

    st.subheader("2 · Upload your spreadsheet")
    upload = st.file_uploader("CSV or Excel file", type=["csv", "xlsx", "xls"], key="bulk_upload")
    if upload is None:
        st.info("Drop a CSV or Excel file to begin.")
        return

    try:
        df = read_table(upload, upload.name)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read the file: {exc}")
        return

    if df.empty:
        st.warning("That file has no data rows.")
        return

    st.success(f"Loaded **{len(df)}** rows and {len(df.columns)} columns.")
    with st.expander("👁️ Preview first rows", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)

    # ---- Column mapping ------------------------------------------------- #
    st.subheader("3 · Map columns")
    st.caption("Match each XML tag to a column from your file. Guesses are pre-filled.")
    columns = list(df.columns)
    options = ["— none —"] + columns
    mapping = {}

    map_cols = st.columns(2)
    targets = [(f["name"], f["label"], f["required"]) for f in FIELDS] + [
        (DESCRIPTION_FIELD, "Job Description", True)
    ]
    for i, (name, label, required) in enumerate(targets):
        guess = guess_column(name, columns)
        default_idx = options.index(guess) if guess in options else 0
        sel = map_cols[i % 2].selectbox(
            f"{label}{' *' if required else ''}  →  <{name}>",
            options, index=default_idx, key=f"map_{name}",
        )
        mapping[name] = None if sel == "— none —" else sel

    desc_is_html = st.checkbox(
        "Description column already contains HTML (skip plain-text → HTML conversion)",
        value=False,
        help="Leave unchecked if the description is plain text; the app will build the HTML for you.",
    )

    # ---- Generate ------------------------------------------------------- #
    st.subheader("4 · Generate XML")
    if st.button("⚡ Generate XML from spreadsheet", type="primary", key="gen_bulk"):
        jobs, skipped = build_jobs_from_df(df, mapping, desc_is_html)
        if not jobs:
            st.error("No valid rows. Check your column mapping — every required field must map to a column with data.")
            st.session_state.bulk_xml_bytes = None
        else:
            st.session_state.bulk_xml_bytes = build_xml(
                jobs, wrap_desc_cdata=wrap_desc, wrap_url_cdata=wrap_url,
                publisher=publisher or None, publisher_url=publisher_url or None,
            )
            st.success(f"Generated feed with **{len(jobs)}** job(s).")
            if skipped:
                with st.expander(f"⚠️ {len(skipped)} row(s) skipped (missing required data)"):
                    st.dataframe(
                        [{"Spreadsheet row": r, "Missing fields": ", ".join(m)} for r, m in skipped],
                        use_container_width=True, hide_index=True,
                    )

    if st.session_state.bulk_xml_bytes:
        with st.expander("📄 XML preview", expanded=True):
            preview = st.session_state.bulk_xml_bytes.decode("utf-8")
            st.code(preview[:20000] + ("\n... (truncated preview)" if len(preview) > 20000 else ""),
                    language="xml")
        st.download_button(
            "⬇️ Download .xml",
            data=st.session_state.bulk_xml_bytes,
            file_name=filename,
            mime="application/xml",
            type="primary",
            use_container_width=True,
            key="dl_bulk",
        )


if __name__ == "__main__":
    main()
