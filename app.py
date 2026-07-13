import streamlit as st
import re
import PyPDF2
from docx import Document

st.set_page_config(
    page_title="Offer Letter Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================================
# Accenture-themed styling
#   Brand: Accenture Purple #A100FF, black, white, the ">" mark
# ============================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    :root {
        --acc-purple: #A100FF;
        --acc-purple-dk: #7500C0;
        --acc-magenta: #E619E6;
        --acc-black: #101010;
        --acc-grey: #6E6E80;
    }

    html, body, [class*="css"], .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    [data-testid="stMainBlockContainer"] {
        padding-top: 1.5rem;
        max-width: 1180px;
    }

    /* Hero header */
    .acc-hero {
        background: linear-gradient(120deg, #A100FF 0%, #7500C0 45%, #460073 100%);
        border-radius: 20px;
        padding: 38px 40px;
        color: #fff;
        margin-bottom: 28px;
        box-shadow: 0 18px 45px rgba(161, 0, 255, 0.35);
        position: relative;
        overflow: hidden;
    }
    .acc-hero::after {
        content: ">";
        position: absolute;
        right: 24px;
        top: 50%;
        transform: translateY(-50%);
        font-size: 190px;
        font-weight: 800;
        color: rgba(255,255,255,0.10);
        line-height: 1;
    }
    .acc-hero h1 {
        font-size: 2.5rem;
        font-weight: 900;
        letter-spacing: -1.2px;
        margin: 0 0 6px 0;
        color: #fff;
    }
    .acc-hero h1 .gt { color: #ffffff; opacity: .85; margin-right: 4px; }
    .acc-hero p {
        font-size: 1.05rem;
        opacity: 0.92;
        margin: 0;
        max-width: 640px;
    }
    .acc-badge {
        display: inline-block;
        background: rgba(255,255,255,0.16);
        border: 1px solid rgba(255,255,255,0.28);
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: .4px;
        margin-top: 16px;
    }

    /* Section headings */
    .acc-section {
        font-size: 1.5rem;
        font-weight: 800;
        color: var(--acc-black);
        margin: 2.2rem 0 0.4rem;
        letter-spacing: -0.4px;
    }
    .acc-section .gt { color: var(--acc-purple); margin-right: 8px; font-weight: 900; }
    .acc-sub { color: var(--acc-grey); font-size: 0.95rem; margin-bottom: 1.2rem; }

    /* Card panels */
    .acc-panel-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: var(--acc-purple-dk);
        border-left: 4px solid var(--acc-purple);
        padding-left: 12px;
        margin-bottom: 6px;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(120deg, #A100FF, #7500C0);
        color: #fff;
        border: none;
        border-radius: 12px;
        padding: 0.7rem 1.4rem;
        font-weight: 700;
        font-size: 1rem;
        letter-spacing: .2px;
        box-shadow: 0 8px 22px rgba(161,0,255,0.30);
        transition: transform .15s ease, box-shadow .15s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 28px rgba(161,0,255,0.42);
        color: #fff;
    }
    .stButton > button:focus:not(:active) { color: #fff; }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: #fff;
        border: 1px solid #EFE6FA;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 6px 18px rgba(80,0,120,0.06);
    }
    [data-testid="stMetricValue"] {
        color: var(--acc-purple-dk);
        font-weight: 900;
    }
    [data-testid="stMetricLabel"] {
        color: var(--acc-grey);
        font-weight: 600;
    }

    /* Big hike banner */
    .hike-banner {
        border-radius: 20px;
        padding: 34px;
        text-align: center;
        margin: 8px 0 20px;
        color: #fff;
    }
    .hike-up   { background: linear-gradient(120deg, #A100FF, #7500C0); box-shadow: 0 16px 40px rgba(161,0,255,.32); }
    .hike-down { background: linear-gradient(120deg, #6E6E80, #3a3a44); box-shadow: 0 16px 40px rgba(60,60,70,.32); }
    .hike-banner .lbl { font-size: .85rem; letter-spacing: 2px; text-transform: uppercase; opacity: .9; font-weight: 700; }
    .hike-banner .val { font-size: 4.2rem; font-weight: 900; letter-spacing: -3px; line-height: 1.05; margin: 6px 0; }
    .hike-banner .dlt { font-size: 1.05rem; opacity: .95; }
    .hike-banner .dlt b { font-weight: 800; }

    /* Dataframe */
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }

    /* Uploader */
    [data-testid="stFileUploaderDropzone"] {
        background: #FBF7FF;
        border: 2px dashed #D9BFF5;
        border-radius: 14px;
    }

    /* Progress bar color */
    [data-testid="stProgress"] > div > div > div > div {
        background: linear-gradient(90deg, #A100FF, #E619E6);
    }

    .fineprint { color: var(--acc-grey); font-size: 0.82rem; text-align: center; margin-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# Parsing logic
# ============================================================================

CURRENCIES = {"INR": "₹", "USD": "$", "EUR": "€", "GBP": "£", "AED": "د.إ"}

KEYWORDS = {
    "ctc": r"(total\s+)?(annual\s+)?(ctc|cost\s+to\s+(the\s+)?company|total\s+(annual\s+)?compensation|total\s+(annual\s+)?remuneration|annual\s+(gross\s+)?(salary|package|compensation)|gross\s+annual|total\s+pay|package)",
    "base": r"(base|basic)\s+(salary|pay|component)|fixed\s+(pay|salary|compensation|component)|annual\s+base",
    "variable": r"variable\s+(pay|component|compensation|bonus)|performance\s+(bonus|pay|incentive)|annual\s+bonus|target\s+bonus|incentive\s+pay",
    "joining": r"(joining|sign[\s-]?on|signing|welcome)\s+bonus|one[\s-]?time\s+(bonus|payment)",
    "stock": r"rsu|esop|stock|equity|restricted\s+share|share\s+units?|espp\s+grant",
}

MONTHLY_RE = re.compile(r"per\s+month|monthly|p\.?m\.?\b|\/\s*month", re.I)
ANNUAL_RE = re.compile(r"per\s+annum|annual|yearly|p\.?a\.?\b|\/\s*year|lpa", re.I)
AMOUNT_RE = re.compile(
    r"(?:₹|rs\.?|inr|\$|usd|€|eur|£|gbp|aed)?\s*([\d]{1,3}(?:,\d{2,3})*(?:\.\d+)?|\d+(?:\.\d+)?)\s*(lpa|lakhs?|lacs?|crores?|cr\b|k\b|million|mn\b)?",
    re.I,
)


def to_number(num_str, unit=""):
    try:
        n = float(num_str.replace(",", ""))
    except ValueError:
        return None
    unit = (unit or "").lower()
    if re.search(r"lpa|lakh|lac", unit):
        n *= 100000
    elif re.search(r"crore|cr\b", unit):
        n *= 10000000
    elif re.search(r"^k$", unit):
        n *= 1000
    elif re.search(r"million|mn|m$", unit):
        n *= 1000000
    return int(round(n))


def amounts_in(text):
    amounts = []
    for m in AMOUNT_RE.finditer(text):
        raw, unit = m.group(1), m.group(2)
        val = to_number(raw, unit)
        if val is None:
            continue
        has_unit = bool(unit)
        has_symbol = bool(re.search(r"₹|rs\.?|inr|\$|usd|€|eur|£|gbp|aed", m.group(0), re.I))
        has_commas = "," in raw
        if not has_unit and not has_symbol and not has_commas and val < 10000:
            continue
        if val < 100:
            continue
        amounts.append(val)
    return amounts


def extract_field(lines, key):
    pattern = re.compile(KEYWORDS[key], re.I)
    candidates = []
    for i, line in enumerate(lines):
        if not pattern.search(line):
            continue
        search_text = line
        vals = amounts_in(search_text)
        if not vals and i + 1 < len(lines):
            search_text = line + " " + lines[i + 1]
            vals = amounts_in(lines[i + 1])
        if not vals:
            continue
        v = max(vals)
        if MONTHLY_RE.search(search_text) and not ANNUAL_RE.search(search_text):
            v *= 12
        candidates.append(v)
    if not candidates:
        return None
    return max(candidates) if key == "ctc" else candidates[0]


def extract_offer(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    result = {k: extract_field(lines, k) for k in ["ctc", "base", "variable", "joining", "stock"]}
    if result["ctc"] and result["base"] and result["ctc"] < result["base"]:
        result["ctc"], result["base"] = result["base"], result["ctc"]
    return result


def detect_currency(text):
    scores = {
        "INR": len(re.findall(r"₹|(?:^|\W)(rs\.?|inr)(?:\W)|lpa|lakh|lac|crore", text, re.I)),
        "USD": len(re.findall(r"\$|usd", text, re.I)),
        "EUR": len(re.findall(r"€|eur(?:o|\b)", text, re.I)),
        "GBP": len(re.findall(r"£|gbp", text, re.I)),
        "AED": len(re.findall(r"aed|dirham", text, re.I)),
    }
    top = max(scores, key=scores.get)
    return top if scores[top] > 0 else "INR"


def read_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        reader = PyPDF2.PdfReader(uploaded_file)
        text = ""
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
        if len(text.strip()) < 30:
            raise ValueError("This PDF looks scanned (no selectable text). Please paste the text instead.")
        return text
    if name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    return uploaded_file.getvalue().decode("utf-8", errors="ignore")


def fmt_money(n, currency):
    if n is None:
        return "—"
    n = int(n)
    sym = CURRENCIES.get(currency, "")
    if currency == "INR":
        if n >= 10000000:
            return sym + (f"{n / 10000000:.2f}".rstrip("0").rstrip(".")) + " Cr"
        if n >= 100000:
            return sym + (f"{n / 100000:.2f}".rstrip("0").rstrip(".")) + " L"
        return f"{sym}{n:,}"
    if n >= 1000000:
        return sym + (f"{n / 1000000:.1f}".rstrip("0").rstrip(".")) + "M"
    if n >= 1000:
        return sym + (f"{n / 1000:.1f}".rstrip("0").rstrip(".")) + "K"
    return f"{sym}{n:,}"


def fmt_full(n, currency):
    if n is None:
        return "—"
    return f"{CURRENCIES.get(currency, '')}{int(n):,}"


FIELDS = ["ctc", "base", "variable", "joining", "stock"]
FIELD_LABELS = {
    "ctc": "Total CTC / Annual Package",
    "base": "Base / Fixed Salary",
    "variable": "Variable Pay / Performance Bonus",
    "joining": "Joining / Sign-on Bonus",
    "stock": "Stocks / RSU / ESOP (per year)",
}

# ============================================================================
# UI
# ============================================================================

st.markdown("""
<div class="acc-hero">
    <h1><span class="gt">&gt;</span>Offer Letter Analyzer</h1>
    <p>Upload your previous and current offer letters to instantly compare salary, hike percentage, bonuses and benefits — clearly and confidently.</p>
    <div class="acc-badge">🔒 Private — files are processed in this session and never stored</div>
</div>
""", unsafe_allow_html=True)

# ---- Session state init ----
for flag in ["extracted", "calculated"]:
    if flag not in st.session_state:
        st.session_state[flag] = False

# ---- Step 1: Upload ----
st.markdown('<div class="acc-section"><span class="gt">&gt;</span>Step 1 · Upload Offer Letters</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="acc-panel-title">Previous Org Offer</div>', unsafe_allow_html=True)
    prev_file = st.file_uploader("Upload PDF, DOCX or TXT", key="prev_file", type=["pdf", "docx", "txt"])
    prev_text = st.text_area("…or paste the text here", key="prev_text", height=140,
                             placeholder="Paste your previous offer letter text…") if not prev_file else None
with c2:
    st.markdown('<div class="acc-panel-title">Current Offer</div>', unsafe_allow_html=True)
    curr_file = st.file_uploader("Upload PDF, DOCX or TXT", key="curr_file", type=["pdf", "docx", "txt"])
    curr_text = st.text_area("…or paste the text here", key="curr_text", height=140,
                             placeholder="Paste your current offer letter text…") if not curr_file else None

if st.button("🔍  Read & Extract Data", width='stretch'):
    try:
        prev_content = read_file(prev_file) if prev_file else prev_text
        curr_content = read_file(curr_file) if curr_file else curr_text
        if not prev_content or not str(prev_content).strip():
            st.error("Please provide the previous offer letter (upload or paste).")
            st.stop()
        if not curr_content or not str(curr_content).strip():
            st.error("Please provide the current offer letter (upload or paste).")
            st.stop()

        prev_data = extract_offer(prev_content)
        curr_data = extract_offer(curr_content)
        currency = detect_currency(prev_content + " " + curr_content)

        # Seed widget values via session_state so number inputs pre-fill
        for f in FIELDS:
            st.session_state[f"prev_{f}"] = int(prev_data[f] or 0)
            st.session_state[f"curr_{f}"] = int(curr_data[f] or 0)
        st.session_state["currency_sel"] = currency
        st.session_state["extracted"] = True
        st.session_state["calculated"] = False

        found = sum(1 for f in FIELDS if prev_data[f]) + sum(1 for f in FIELDS if curr_data[f])
        st.session_state["found_count"] = found
    except Exception as e:
        st.error(f"❌ Could not read file: {e}")

# ---- Step 2: Review ----
if st.session_state["extracted"]:
    st.markdown('<div class="acc-section"><span class="gt">&gt;</span>Step 2 · Review What We Found</div>', unsafe_allow_html=True)
    fc = st.session_state.get("found_count", 0)
    if fc == 0:
        st.markdown('<div class="acc-sub">No amounts were auto-detected — please enter the figures manually below. Enter <b>annual</b> values.</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="acc-sub">Auto-detected <b>{fc}</b> value(s). Please verify and correct anything below. Enter <b>annual</b> values.</div>', unsafe_allow_html=True)

    cur_col, _ = st.columns([1, 2])
    with cur_col:
        currency_sel = st.selectbox("Currency", list(CURRENCIES.keys()), key="currency_sel")
    step = 100000 if currency_sel == "INR" else 1000

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="acc-panel-title">Previous Offer</div>', unsafe_allow_html=True)
        for f in FIELDS:
            st.number_input(FIELD_LABELS[f], min_value=0, step=step, key=f"prev_{f}")
    with col2:
        st.markdown('<div class="acc-panel-title">Current Offer</div>', unsafe_allow_html=True)
        for f in FIELDS:
            st.number_input(FIELD_LABELS[f], min_value=0, step=step, key=f"curr_{f}")

    if st.button("⚡  Calculate Hike & Comparison", width='stretch'):
        st.session_state["calculated"] = True

# ---- Step 3: Results ----
if st.session_state["extracted"] and st.session_state["calculated"]:
    currency = st.session_state["currency_sel"]
    prev = {f: int(st.session_state[f"prev_{f}"]) for f in FIELDS}
    curr = {f: int(st.session_state[f"curr_{f}"]) for f in FIELDS}

    prev_total = prev["ctc"] or (prev["base"] + prev["variable"])
    curr_total = curr["ctc"] or (curr["base"] + curr["variable"])

    if not prev_total or not curr_total:
        st.error("Please fill in Total CTC (or at least Base salary) for both offers.")
        st.stop()

    hike = (curr_total - prev_total) / prev_total * 100
    diff = curr_total - prev_total
    monthly_diff = diff / 12

    st.markdown('<div class="acc-section"><span class="gt">&gt;</span>Step 3 · Your Results</div>', unsafe_allow_html=True)

    cls = "hike-up" if hike >= 0 else "hike-down"
    sign = "+" if hike >= 0 else ""
    st.markdown(f"""
    <div class="hike-banner {cls}">
        <div class="lbl">Overall Hike</div>
        <div class="val">{sign}{hike:.1f}%</div>
        <div class="dlt">That's <b>{fmt_full(abs(diff), currency)}</b> {'more' if diff >= 0 else 'less'} per year ·
        <b>{fmt_full(abs(monthly_diff), currency)}</b> per month (gross)</div>
    </div>
    """, unsafe_allow_html=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Previous Total CTC", fmt_money(prev_total, currency), fmt_full(prev_total, currency))
    m2.metric("Current Total CTC", fmt_money(curr_total, currency), fmt_full(curr_total, currency))
    m3.metric("Monthly Gross (New)", fmt_money(curr_total / 12, currency), "≈ CTC ÷ 12, before tax")

    st.markdown("&nbsp;", unsafe_allow_html=True)

    # CTC comparison bars
    st.markdown('<div class="acc-panel-title">Total CTC — Previous vs Current</div>', unsafe_allow_html=True)
    b1, b2 = st.columns(2)
    max_v = max(prev_total, curr_total)
    with b1:
        st.caption(f"Previous · {fmt_full(prev_total, currency)}")
        st.progress(prev_total / max_v)
    with b2:
        st.caption(f"Current · {fmt_full(curr_total, currency)}")
        st.progress(curr_total / max_v)

    st.markdown("&nbsp;", unsafe_allow_html=True)

    # Component breakdown
    st.markdown('<div class="acc-panel-title">Component-wise Breakdown</div>', unsafe_allow_html=True)
    rows = []
    for f in FIELDS:
        p, c = prev[f], curr[f]
        if not p and not c:
            continue
        d = c - p
        pct = f" ({d / p * 100:+.1f}%)" if p > 0 else ""
        arrow = "📈" if d > 0 else "📉" if d < 0 else "➡️"
        rows.append({
            "Component": FIELD_LABELS[f],
            "Previous": fmt_full(p, currency) if p else "—",
            "Current": fmt_full(c, currency) if c else "—",
            "Change": (f"{'+' if d >= 0 else '−'}{fmt_full(abs(d), currency)}{pct} {arrow}") if d != 0 else "—",
        })
    rows.append({
        "Component": "Monthly Gross (÷ 12)",
        "Previous": fmt_full(prev_total / 12, currency),
        "Current": fmt_full(curr_total / 12, currency),
        "Change": f"{'+' if monthly_diff >= 0 else '−'}{fmt_full(abs(monthly_diff), currency)} 📈",
    })
    st.dataframe(rows, width='stretch', hide_index=True)

    # Takeaways
    st.markdown('<div class="acc-panel-title">Key Takeaways</div>', unsafe_allow_html=True)
    t1, t2, t3 = st.columns(3)
    t1.info(f"**Annual increase:** {fmt_full(diff, currency)}")
    if curr["stock"] > prev["stock"] and curr["stock"] > 0:
        t2.success(f"**New/higher Stock:** {fmt_full(curr['stock'], currency)}/yr 🎉")
    elif curr["joining"]:
        t2.success(f"**Joining bonus:** {fmt_full(curr['joining'], currency)}")
    else:
        t2.info("**Bonus:** review the breakdown above")
    t3.warning(f"**New monthly gross:** {fmt_full(curr_total / 12, currency)}")

    st.markdown('<div class="fineprint">Figures are parsed automatically from your documents — always cross-check with the original offer letters. Monthly ≈ annual ÷ 12 (gross, before tax).</div>', unsafe_allow_html=True)
