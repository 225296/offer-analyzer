import streamlit as st
import re
from io import BytesIO
import PyPDF2
from docx import Document

st.set_page_config(page_title="Offer Letter Analyzer", layout="wide", initial_sidebar_state="collapsed")

# ============================================================================
# CSS Styling
# ============================================================================
st.markdown("""
<style>
    [data-testid="stMainBlockContainer"] {
        padding-top: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }
    .stat-value {
        font-size: 2.5em;
        font-weight: 900;
        margin: 10px 0;
    }
    .stat-label {
        font-size: 0.9em;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .hike-up { color: #34d399; }
    .hike-down { color: #f87171; }
    .section-header {
        font-size: 1.8em;
        font-weight: 800;
        margin: 2rem 0 1rem;
        border-bottom: 2px solid #667eea;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# Parse Functions
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
    re.I
)


def to_number(num_str, unit=""):
    """Convert a matched amount string to a number."""
    try:
        n = float(num_str.replace(",", ""))
    except:
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
    """Find all money-like values in text."""
    amounts = []
    for m in AMOUNT_RE.finditer(text):
        raw = m.group(1)
        unit = m.group(2)
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
    """Extract a specific field from offer letter lines."""
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
    """Extract all offer components from text."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    result = {}

    for key in ["ctc", "base", "variable", "joining", "stock"]:
        result[key] = extract_field(lines, key)

    # Sanity check: CTC should not be less than base
    if result["ctc"] and result["base"] and result["ctc"] < result["base"]:
        result["ctc"], result["base"] = result["base"], result["ctc"]

    return result


def detect_currency(text):
    """Detect currency from text."""
    scores = {
        "INR": len(re.findall(r"₹|(?:^|\W)(rs\.?|inr)(?:\W)|lpa|lakh|lac|crore", text, re.I)),
        "USD": len(re.findall(r"\$|usd", text, re.I)),
        "EUR": len(re.findall(r"€|eur(?:o|\b)", text, re.I)),
        "GBP": len(re.findall(r"£|gbp", text, re.I)),
        "AED": len(re.findall(r"aed|dirham", text, re.I)),
    }
    top_currency = max(scores, key=scores.get)
    return top_currency if scores[top_currency] > 0 else "INR"


def read_file(uploaded_file):
    """Read content from uploaded file."""
    if uploaded_file.type == "application/pdf":
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return text

    elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs])
        return text

    else:  # txt
        return uploaded_file.getvalue().decode("utf-8")


def fmt_money(n, currency):
    """Format number as money."""
    if n is None or (isinstance(n, float) and n != n):  # NaN check
        return "—"

    n = int(n)
    sym = CURRENCIES.get(currency, "")

    if currency == "INR":
        if n >= 10000000:
            return f"{sym}{n / 10000000:.2f}".rstrip("0").rstrip(".") + " Cr"
        if n >= 100000:
            return f"{sym}{n / 100000:.2f}".rstrip("0").rstrip(".") + " L"
        return f"{sym}{n:,}"

    if n >= 1000000:
        return f"{sym}{n / 1000000:.1f}".rstrip("0").rstrip(".") + "M"
    if n >= 1000:
        return f"{sym}{n / 1000:.1f}".rstrip("0").rstrip(".") + "K"

    return f"{sym}{n:,}"


def fmt_full(n, currency):
    """Format number as full currency amount."""
    if n is None or (isinstance(n, float) and n != n):
        return "—"

    n = int(n)
    sym = CURRENCIES.get(currency, "")
    return f"{sym}{n:,}"


# ============================================================================
# Main UI
# ============================================================================

st.markdown("# 📊 Offer Letter Analyzer")
st.markdown("Upload your previous and current offer letters to compare salary, hike %, bonus, and benefits.")

st.markdown('<div style="background: rgba(52,211,153,0.1); border: 1px solid rgba(52,211,153,0.3); padding: 12px; border-radius: 8px; margin-bottom: 2rem;"><b>🔒 Privacy:</b> Files are processed in your browser session and never stored.</div>', unsafe_allow_html=True)

# ============================================================================
# File Upload
# ============================================================================

col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Previous Org Offer")
    prev_file = st.file_uploader("Upload or paste previous offer", key="prev", type=["pdf", "docx", "txt"])

    if not prev_file:
        prev_text = st.text_area("Or paste the text here:", key="prev_text", height=200, placeholder="Paste your previous offer letter text...")
    else:
        prev_text = None

with col2:
    st.subheader("📄 Current Org Offer")
    curr_file = st.file_uploader("Upload or paste current offer", key="curr", type=["pdf", "docx", "txt"])

    if not curr_file:
        curr_text = st.text_area("Or paste the text here:", key="curr_text", height=200, placeholder="Paste your current offer letter text...")
    else:
        curr_text = None

# Extract text from files or use pasted text
prev_content = None
curr_content = None
extracted = False

if st.button("🔍 Read & Extract Data", type="primary", use_container_width=True):
    try:
        if prev_file:
            prev_content = read_file(prev_file)
        elif prev_text:
            prev_content = prev_text
        else:
            st.error("Please upload or paste your previous offer letter")
            st.stop()

        if curr_file:
            curr_content = read_file(curr_file)
        elif curr_text:
            curr_content = curr_text
        else:
            st.error("Please upload or paste your current offer letter")
            st.stop()

        extracted = True
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        st.stop()

if extracted:
    # Extract data
    prev_data = extract_offer(prev_content)
    curr_data = extract_offer(curr_content)
    detected_currency = detect_currency(prev_content + " " + curr_content)

    # ============================================================================
    # Step 2: Review and Edit
    # ============================================================================

    st.markdown("<div class='section-header'>Step 2 · Review Extracted Data</div>", unsafe_allow_html=True)
    st.markdown("Verify the auto-detected values below. Edit anything that looks wrong. Enter **annual** figures.")

    col_cur, _ = st.columns([2, 3])
    with col_cur:
        currency_sel = st.selectbox("Select Currency:", list(CURRENCIES.keys()), index=list(CURRENCIES.keys()).index(detected_currency))

    col1, col2 = st.columns(2)

    fields_to_extract = ["ctc", "base", "variable", "joining", "stock"]
    field_labels = {
        "ctc": "Total CTC / Annual Package",
        "base": "Base / Fixed Salary",
        "variable": "Variable Pay / Performance Bonus",
        "joining": "Joining / Sign-on Bonus",
        "stock": "Stocks / RSU / ESOP (per year)",
    }

    with col1:
        st.subheader("Previous Offer")
        prev_edited = {}
        for field in fields_to_extract:
            val = prev_data[field]
            prev_edited[field] = st.number_input(
                field_labels[field],
                value=val if val else 0.0,
                min_value=0.0,
                step=100000.0 if currency_sel == "INR" else 1000.0,
                key=f"prev_{field}",
            )

    with col2:
        st.subheader("Current Offer")
        curr_edited = {}
        for field in fields_to_extract:
            val = curr_data[field]
            curr_edited[field] = st.number_input(
                field_labels[field],
                value=val if val else 0.0,
                min_value=0.0,
                step=100000.0 if currency_sel == "INR" else 1000.0,
                key=f"curr_{field}",
            )

    # ============================================================================
    # Step 3: Calculate
    # ============================================================================

    if st.button("⚡ Calculate Hike & Comparison", type="primary", use_container_width=True):
        prev_total = prev_edited["ctc"] or (prev_edited["base"] + prev_edited["variable"])
        curr_total = curr_edited["ctc"] or (curr_edited["base"] + curr_edited["variable"])

        if not prev_total or not curr_total:
            st.error("Please fill in Total CTC (or Base + Variable) for both offers.")
            st.stop()

        hike_pct = ((curr_total - prev_total) / prev_total) * 100
        diff = curr_total - prev_total
        monthly_diff = diff / 12

        # ============================================================================
        # Results
        # ============================================================================

        st.markdown("<div class='section-header'>Step 3 · Your Results</div>", unsafe_allow_html=True)

        # Hike banner
        hike_color = "hike-up" if hike_pct >= 0 else "hike-down"
        hike_sign = "+" if hike_pct >= 0 else ""

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(
                "Overall Hike",
                f"{hike_sign}{hike_pct:.1f}%",
                f"{fmt_full(abs(diff), currency_sel)} {'more' if diff >= 0 else 'less'} per year"
            )

        with col2:
            st.metric(
                "Monthly Gross (New)",
                fmt_money(curr_total / 12, currency_sel),
                "≈ CTC ÷ 12, before tax"
            )

        with col3:
            st.metric(
                "Monthly Increase",
                fmt_full(monthly_diff, currency_sel),
                f"{(hike_pct):.1f}% hike"
            )

        st.divider()

        # Component comparison
        st.markdown("### 💰 Component-wise Comparison")

        comparison_data = []
        for field in fields_to_extract:
            prev_val = prev_edited[field]
            curr_val = curr_edited[field]

            if prev_val or curr_val:
                delta = curr_val - prev_val
                delta_pct = (delta / prev_val * 100) if prev_val > 0 else 0

                comparison_data.append({
                    "Component": field_labels[field],
                    "Previous": fmt_full(prev_val, currency_sel) if prev_val else "—",
                    "Current": fmt_full(curr_val, currency_sel) if curr_val else "—",
                    "Change": f"{fmt_full(abs(delta), currency_sel)} {'📈' if delta > 0 else '📉' if delta < 0 else '➡️'}" + (f" ({delta_pct:+.1f}%)" if prev_val > 0 else ""),
                })

        df_comparison = st.dataframe(
            comparison_data,
            use_container_width=True,
            hide_index=True,
        )

        st.divider()

        # Bars
        st.markdown("### 📊 Total CTC Comparison")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Previous Total CTC", fmt_full(prev_total, currency_sel))
            st.progress(min(prev_total / max(prev_total, curr_total), 1.0))

        with col2:
            st.metric("Current Total CTC", fmt_full(curr_total, currency_sel))
            st.progress(min(curr_total / max(prev_total, curr_total), 1.0))

        st.divider()

        # Key takeaways
        st.markdown("### 🎯 Key Takeaways")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.info(f"**Annual Increase:** {fmt_full(diff, currency_sel)}")

        with col2:
            if curr_edited["stock"] > prev_edited["stock"]:
                st.success(f"**New Stock/RSU:** {fmt_full(curr_edited['stock'], currency_sel)}/year 🎉")
            elif curr_edited["joining"] > 0:
                st.success(f"**Joining Bonus:** {fmt_full(curr_edited['joining'], currency_sel)}")

        with col3:
            st.warning(f"**Monthly Gross:** {fmt_full(curr_total / 12, currency_sel)}")

        st.divider()
        st.caption("💡 Figures are parsed from your offer letters. Always cross-check with the originals. Monthly = annual ÷ 12 (gross, before tax).")
