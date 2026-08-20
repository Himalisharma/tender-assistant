import streamlit as st
from groq import Groq
from dotenv import load_dotenv
import os
import fitz
import re
import time
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ══════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════
load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY")

st.set_page_config(page_title="Tender Analyzer", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .stApp{background:#f5f6fa}
    section[data-testid="stSidebar"]{background:#1a1a2e!important}
    section[data-testid="stSidebar"] *{color:#e0e0e0!important}
    .block-container *{color:#1a1a2e}
    .page-title{font-size:1.9rem;font-weight:700;color:#1a1a2e!important;margin-bottom:.2rem}
    .page-sub{font-size:1rem;color:#555!important;margin-bottom:1.5rem}
    .info-box{background:#e8eaf6;border-left:4px solid #3f51b5;padding:.8rem 1rem;
        border-radius:0 10px 10px 0;font-size:.88rem;margin-bottom:1rem}
    .success-box{background:#e8f5e9;border-left:4px solid #43a047;padding:.8rem 1rem;
        border-radius:0 10px 10px 0;font-size:.88rem;color:#1b5e20!important;margin-bottom:1rem}
    .warn-box{background:#fff8e1;border-left:4px solid #ffa000;padding:.8rem 1rem;
        border-radius:0 10px 10px 0;font-size:.88rem;margin-bottom:1rem}
    .chat-user-wrap{display:flex;justify-content:flex-end;margin:.5rem 0}
    .chat-bot-wrap{display:flex;justify-content:flex-start;margin:.5rem 0;align-items:flex-start}
    .chat-user{background:#3f51b5;color:white!important;padding:.7rem 1.1rem;
        border-radius:18px 18px 4px 18px;max-width:75%;font-size:.93rem;line-height:1.55}
    .chat-bot{background:white;color:#1a1a2e!important;padding:.7rem 1.1rem;
        border-radius:18px 18px 18px 4px;max-width:80%;font-size:.93rem;
        line-height:1.65;border:1px solid #e8eaf6}
    .bot-avatar{width:32px;height:32px;border-radius:50%;background:#e8eaf6;
        display:flex;align-items:center;justify-content:center;font-size:1rem;
        margin-right:8px;flex-shrink:0;margin-top:2px}
    .preview-table{width:100%;border-collapse:collapse;font-size:.85rem}
    .preview-table th{background:#1a1a2e;color:white;padding:8px 12px;text-align:left}
    .preview-table td{padding:8px 12px;border-bottom:1px solid #f0f0f0;vertical-align:top}
    .preview-table tr:nth-child(even) td{background:#fafafa}
    .found{color:#1b5e20;font-weight:500}
    .missing{color:#aaa;font-style:italic}
    #MainMenu,footer,header{visibility:hidden}
    div[data-testid="stButton"]>button{border-radius:10px!important;font-weight:600!important}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
# PDF READING
# ══════════════════════════════════════════════════════════════════════════

def read_pdf(uploaded_file):
    data  = uploaded_file.read()
    doc   = fitz.open(stream=data, filetype="pdf")
    lines = []
    full  = ""
    for i, page in enumerate(doc):
        t = page.get_text("text")
        full += t + "\n"
        for ln in t.split("\n"):
            ln = ln.strip()
            if ln:
                lines.append((i + 1, ln))
    doc.close()
    uploaded_file.seek(0)
    return lines, full


def clean(text):
    for bad, good in {'\ufb01':'fi','\ufb02':'fl','\u00ad':'','\u200b':''}.items():
        text = text.replace(bad, good)
    return re.sub(r'\n{3,}', '\n\n', re.sub(r' {3,}', ' ', text))


def search_pdf(lines, keywords, ctx=35, max_s=5):
    """Find the most relevant sections in the PDF for given keywords."""
    texts = [clean(ln) for _, ln in lines]
    hits  = {}
    for kw in keywords:
        kw = kw.lower()
        for i, ln in enumerate(texts):
            if kw in ln.lower():
                s = max(0, i - 3)
                e = min(len(texts), i + ctx)
                if s not in hits:
                    hits[s] = {"score": 0, "end": e}
                hits[s]["score"] += 1
                hits[s]["end"]    = max(hits[s]["end"], e)
    if not hits:
        return ""
    snippets, covered = [], set()
    for s in sorted(hits, key=lambda x: -hits[x]["score"]):
        e = hits[s]["end"]
        if len(covered & set(range(s, e))) > (e - s) * 0.6:
            continue
        snippets.append("\n".join(texts[s:e]))
        covered.update(range(s, e))
        if len(snippets) >= max_s:
            break
    return "\n\n---\n\n".join(snippets)


# ══════════════════════════════════════════════════════════════════════════
# GROQ API CALL
# ══════════════════════════════════════════════════════════════════════════

def call_groq(client, prompt, max_tokens=600):
    for attempt in range(4):
        try:
            r = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content":
                     "You are a government tender document analyst. "
                     "When asked to extract a field, return ONLY the actual value. "
                     "Never repeat the field name. Never copy prompt instructions. "
                     "If the information is not present, say exactly: Not specified in document"},
                    {"role": "user", "content": prompt},
                ],
                max_tokens=max_tokens,
                temperature=0.0,
            )
            return r.choices[0].message.content.strip()
        except Exception as e:
            if "429" in str(e) or "rate" in str(e).lower():
                time.sleep(5 * (attempt + 1))
            else:
                raise
    return ""



# ══════════════════════════════════════════════════════════════════════════
# UNIVERSAL SECTIONS — works for ANY government tender
# ══════════════════════════════════════════════════════════════════════════

SECTIONS = [
    {
        "id":    "overview",
        "title": "Tender Overview",
        "fields": [
            ("Issuing Organization",      ["issued by","purchaser","tendering authority","company","department","ministry","corporation"]),
            ("Tender No / NIT No",        ["tender no","nit no","rfp no","e-tender","bid no","reference no","tender number","nit number"]),
            ("GeM Bid Number",            ["gem bid","gem no","gemc","government e-marketplace","gem portal","gem bid number"]),
            ("Bid Submission Deadline",   ["bid closing","last date","closing date","submission date","due date","bid end date"]),
            ("EMD / Bid Security",        ["earnest money","emd","bid security","bid bond","emd amount"]),
            ("Offer Validity",            ["offer validity","validity period","bid validity","120 days"]),
            ("Pre-Bid Meeting",           ["pre bid","pre-bid meeting","pre bid meeting","details mentioned"]),
            ("Integrity Pact",            ["integrity pact","integrity"]),
        ]
    },
    {
        "id":    "scope",
        "title": "Scope of Work",
        "fields": [
            ("Project Background / Introduction", ["introduction","background","about","subsidiary","operates","refinery","digital transformation"]),
            ("Project Objectives",               ["objective","objectives","primary objective","aim","purpose","2.1","2.2","digitizing"]),
            ("MRPL / Buyer Scope",               ["mrpl scope","3.","buyer scope","purchaser scope","sharing","ensuring","ofc"]),
            ("Bidder Scope of Supply",           ["scope of supply","4.1","supply","hand held","tablets","software","networking","hardware"]),
            ("Bidder Scope of Service",          ["scope of service","4.2","installation","commissioning","execution","bidder scope includes"]),
            ("Area / Location of Work",          ["location","area","site","district","state","place of work","mangalore","karnataka","refinery"]),
            ("Package Summary (Nutshell)",       ["scope","objective","digital log","digitiz","overview","implement","entire refinery"]),
        ]
    },
    {
        "id":    "commercial",
        "title": "Commercial Terms",
        "fields": [
            ("Delivery / Completion Period",    ["delivery period","completion period","days from","delivery schedule","delivery / completion","line item 1","line item 2"]),
            ("Payment Terms",                  ["payment term","payment schedule","milestone","11. payment","terms of payment","perpetual type","subscription type"]),
            ("Warranty Period",                ["warranty","warranty period","7. warranty","7.1","7.2","dlp","defect liability","5 year","10 year"]),
            ("CPBG / Performance Security",    ["cpbg","performance bank","pbg","performance security","security deposit","epbg","5% of order"]),
            ("Price Reduction Schedule (PRS)", ["price reduction","prs","0.5%","delay","liquidated","penalty"]),
            ("Contract Duration",              ["contract duration","18 months","contract period","total contract","stage-1","stage-2"]),
            ("Testing & Validation",           ["testing","validation","fat","sat","factory acceptance","site acceptance","5.1","5.2"]),
            ("Training Requirements",          ["training","8. training","train","classroom","field training","8.1","8.2"]),
            ("Documentation Requirements",     ["documentation","9.","system architecture","bill of materials","manuals","as built"]),
        ]
    },
    {
        "id":    "pqc",
        "title": "Pre-Qualification Criteria",
        "fields": [
            ("Prior Experience / Similar Work",  ["similar work","prior experience","executed","refineries","petrochemicals","oil & gas","digital log","10 processing"]),
            ("Financial Turnover Requirement",   ["annual turnover","financial turnover","average annual","crore","turnover","3.5 crore","last 3 years"]),
            ("OEM / Authorization Requirement",  ["oem","authorized","authorization letter","channel partner","manufacturer","group company"]),
            ("OEM Undertaking Requirements",     ["oem undertaking","back to back","guarantee","oem shall assume","complete responsibilities"]),
            ("Net Worth Requirement",            ["net worth","positive net worth","audited","balance sheet","profit & loss"]),
            ("Blacklisting / Holiday List",      ["blacklist","holiday list","debarred","mopng","not under blacklist"]),
            ("MSE / Startup Relaxation",         ["mse","msme","startup","micro and small","relaxation","15%","udyam","nsic"]),
            ("Document Certification Rules",     ["udin","chartered accountant","notarized","statutory auditor","duly certified"]),
        ]
    },
    {
        "id":    "bec",
        "title": "Bid Evaluation Criteria",
        "fields": [
            ("Evaluation Method",              ["l1","lowest","overall l1","techno-commercially","evaluation","landed cost","quantity wise"]),
            ("Technical Track Record",         ["track record","proven","satisfactory performance","performance certificate","working satisfactorily"]),
            ("OEM Support Undertaking",        ["end of support","10 years","oem support","back to back","continuing guarantee"]),
            ("Technical Support Base in India",["technical support","support base","qualified technical","india support","commissioning"]),
            ("Cybersecurity Undertaking",      ["data leaks","cybersecurity","iso 27001","cyber secure","free from"]),
            ("Land Border Restriction",        ["land border","dpiit","registration committee","shares land border","border with india"]),
            ("Rejection Criteria",             ["rejection","liable for rejection","summarily rejected","rejected","non-compliance","false declaration"]),
            ("Special Instructions to Bidders",["special instruction","10.","bidder shall note","bidder must","site visit","proof of concept","nda"]),
        ]
    },
    {
        "id":    "items",
        "title": "Item Details & Specifications",
        "fields": [
            ("Line Items Summary",           ["item details","1400000","line item","description","engineering and supply","to be quoted in gem"]),
            ("Item 1 — Hazardous Tablets",   ["1400000867","hand held device","hazardous area","6\" to 7\"","192","intrinsic safe"]),
            ("Item 2 — Safe Area Tablets",   ["1400000868","safe area tablets","60","tablet"]),
            ("Item 3 — Software Application",["1400000869","software","application","digitizing logs","workflow management","data lake"]),
            ("Item 4 — Networking / Servers",["1400000870","networking infrastructure","servers","cyber secure","network components"]),
            ("Execution Charges (GeM)",      ["1054362","1054364","execution","installation","commissioning","stage-1","stage-2"]),
            ("Per Diem / Expert Visit",      ["1054365","per diem","expert visit","10 year warranty","new requirements","post sat"]),
            ("Tablet Specifications",        ["annexure-1","specification","operating system","ram","storage","display","battery","atex","ip 68","rfid","nfc"]),
            ("Recommended Vendor / Make",    ["recommended vendor","vendor list","annexure-7","rittal","cisco","hp","dell","p&f","i-safe"]),
            ("Asset Details",               ["asset details","static equipment","rotary equipment","instruments","6000","4200","tentative"]),
        ]
    },
]


# ══════════════════════════════════════════════════════════════════════════
# FIELD EXTRACTION
# ══════════════════════════════════════════════════════════════════════════


def extract_field(client, field_name, keywords, lines, full_text):
    """Extract one field from the PDF."""

    # Always include page 1 (letterhead, tender no, company name all appear here)
    page1   = clean(full_text[:3000])
    snippet = search_pdf(lines, keywords, ctx=40, max_s=4)

    # Combine page 1 + keyword-matched sections for best coverage
    if snippet and len(snippet.strip()) > 50:
        context = page1 + "\n\n--- RELEVANT SECTIONS ---\n" + snippet[:3000]
    else:
        context = page1 + "\n\n" + clean(full_text[3000:6000])

    # Cap total context
    context = context[:5500]

    prompt = f"""Read the government tender document below and answer this question:

What is the "{field_name}" in this tender?

DOCUMENT:
{context}

Instructions:
- Give only the direct answer. Example: if asked for "Issuing Organization", just write the company name.
- Do not write the question or field name in your answer.
- Include exact figures, dates, amounts, percentages as written in the document.
- For lists (criteria, milestones, items), number each point clearly.
- If this information is genuinely not in the document, write: Not specified in document

Answer:"""

    result = call_groq(client, prompt, max_tokens=700)

    # Catch echo patterns — LLM repeating the field name or prompt
    echo_signals = [
        f"field: {field_name.lower()}",
        f'"{field_name.lower()}"',
        "the value for",
        "the answer is",
        "in this tender",
        "the field",
        "as requested",
        "based on the document",
        "based on the tender",
        "i found",
        "i was unable",
        "i cannot",
    ]
    result_lower = result.lower()
    if any(sig in result_lower[:80] for sig in echo_signals):
        # LLM echoed the prompt — retry with simpler prompt
        simple_prompt = f"""From this tender document, extract: {field_name}

{context[:3000]}

Just the value, nothing else:"""
        result = call_groq(client, simple_prompt, max_tokens=500)

    # Normalize not-found responses
    not_found = ["not specified in document", "not found", "not mentioned",
                 "not available", "not provided", "n/a", "none", "nil",
                 "not stated", "not present", "not explicitly"]
    if not result or any(result.lower().strip().startswith(p) for p in not_found):
        return "Not specified in document"

    return result


def extract_all(client, lines, full_text, bar, status_ph):
    """Extract all sections and fields."""
    all_results  = {}
    total_fields = sum(len(s["fields"]) for s in SECTIONS)
    done         = 0

    for section in SECTIONS:
        all_results[section["id"]] = {}
        total_sec = len(section["fields"])

        for idx, (field_name, keywords) in enumerate(section["fields"]):
            status_ph.markdown(
                f"<small style='color:#777'>📋 <b>{section['title']}</b> → "
                f"<b>{field_name}</b> ({done+1}/{total_fields})</small>",
                unsafe_allow_html=True
            )
            pct = int((done / max(total_fields, 1)) * 92)
            bar.progress(pct, text=f"Extracting: {field_name[:55]}...")

            if done > 0:
                time.sleep(0.3)

            all_results[section["id"]][field_name] = extract_field(
                client, field_name, keywords, lines, full_text
            )
            done += 1

    return all_results


# ══════════════════════════════════════════════════════════════════════════
# EXCEL GENERATION — creates fresh Excel from extracted data
# ══════════════════════════════════════════════════════════════════════════

def make_excel(all_results, pdf_name):
    wb = Workbook()
    wb.remove(wb.active)

    DARK   = "1A1A2E"
    BLUE   = "3F51B5"
    LBLUE  = "E8EAF6"
    GREEN  = "E8F5E9"
    GFONT  = "1B5E20"
    GREY   = "888888"
    WHITE  = "FFFFFF"

    thin   = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Summary sheet (first) ────────────────────────────────────────────
    ws0 = wb.create_sheet("Summary", 0)

    ws0.merge_cells("A1:C1")
    ws0["A1"] = f"TENDER ANALYSIS REPORT — {pdf_name}"
    ws0["A1"].font      = Font(name="Arial", bold=True, color=WHITE, size=14)
    ws0["A1"].fill      = PatternFill("solid", start_color=DARK)
    ws0["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws0.row_dimensions[1].height = 36

    ws0.merge_cells("A2:C2")
    ws0["A2"] = f"Generated: {datetime.now().strftime('%d %B %Y at %I:%M %p')}"
    ws0["A2"].font      = Font(name="Arial", color=GREY, italic=True, size=9)
    ws0["A2"].alignment = Alignment(horizontal="center")
    ws0.row_dimensions[2].height = 16

    # Header row
    for col, hdr in [(1,"Field"),(2,"Extracted Value"),(3,"Section")]:
        c = ws0.cell(row=4, column=col, value=hdr)
        c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
        c.fill = PatternFill("solid", start_color=BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws0.row_dimensions[4].height = 22

    key_fields = [
        ("Issuing Organization",         "overview"),
        ("Tender No / NIT No",           "overview"),
        ("GeM Bid Number",               "overview"),
        ("Bid Submission Deadline",      "overview"),
        ("EMD / Bid Security",           "overview"),
        ("Offer Validity",               "overview"),
        ("Delivery / Completion Period", "commercial"),
        ("Payment Terms",               "commercial"),
        ("Warranty Period",             "commercial"),
        ("CPBG / Performance Security", "commercial"),
        ("Price Reduction Schedule (PRS)", "commercial"),
        ("Contract Duration",           "commercial"),
        ("Package Summary (Nutshell)",  "scope"),
        ("Area / Location of Work",     "scope"),
        ("Evaluation Method",           "bec"),
        ("Rejection Criteria",          "bec"),
    ]

    for i, (fname, sid) in enumerate(key_fields):
        val     = all_results.get(sid, {}).get(fname, "Not specified in document")
        missing = "not specified" in val.lower()
        r       = 5 + i
        sec_ttl = next((s["title"] for s in SECTIONS if s["id"] == sid), sid)

        lc = ws0.cell(row=r, column=1, value=fname)
        lc.font = Font(name="Arial", bold=True, color=DARK, size=10)
        lc.fill = PatternFill("solid", start_color=LBLUE)
        lc.alignment = Alignment(wrap_text=True, vertical="top")
        lc.border = border

        vc = ws0.cell(row=r, column=2, value=val)
        vc.font = Font(name="Arial", color=GREY if missing else GFONT, italic=missing, size=10)
        vc.fill = PatternFill(fill_type=None) if missing else PatternFill("solid", start_color=GREEN)
        vc.alignment = Alignment(wrap_text=True, vertical="top")
        vc.border = border

        sc = ws0.cell(row=r, column=3, value=sec_ttl)
        sc.font = Font(name="Arial", color=BLUE, size=9)
        sc.alignment = Alignment(vertical="top")
        sc.border = border

        lines_n = max(1, len(val)//80 + val.count("\n"))
        ws0.row_dimensions[r].height = max(20, min(lines_n * 15, 200))

    ws0.column_dimensions["A"].width = 38
    ws0.column_dimensions["B"].width = 70
    ws0.column_dimensions["C"].width = 28
    ws0.freeze_panes = "A5"

    # ── One sheet per section ────────────────────────────────────────────
    for section in SECTIONS:
        sid  = section["id"]
        ws   = wb.create_sheet(title=section["title"][:31])

        # Title row
        ws.merge_cells("A1:C1")
        ws["A1"] = f"{section['title']}  —  {pdf_name}"
        ws["A1"].font      = Font(name="Arial", bold=True, color=WHITE, size=12)
        ws["A1"].fill      = PatternFill("solid", start_color=DARK)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Column headers
        for col, hdr in [(1,"Field"),(2,"Extracted Value from PDF"),(3,"Notes / Manual Input")]:
            c = ws.cell(row=2, column=col, value=hdr)
            c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
            c.fill = PatternFill("solid", start_color=BLUE)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        ws.row_dimensions[2].height = 22

        # Data rows
        for idx, (field_name, _) in enumerate(section["fields"]):
            val     = all_results.get(sid, {}).get(field_name, "Not specified in document")
            missing = "not specified" in val.lower()
            r       = 3 + idx

            lc = ws.cell(row=r, column=1, value=field_name)
            lc.font = Font(name="Arial", bold=True, color=DARK, size=10)
            lc.fill = PatternFill("solid", start_color=LBLUE)
            lc.alignment = Alignment(wrap_text=True, vertical="center")
            lc.border = border

            vc = ws.cell(row=r, column=2, value=val)
            vc.font = Font(name="Arial", color=GREY if missing else "222222",
                           italic=missing, size=10)
            vc.fill = PatternFill(fill_type=None) if missing else PatternFill("solid", start_color=GREEN)
            vc.alignment = Alignment(wrap_text=True, vertical="top")
            vc.border = border

            nc = ws.cell(row=r, column=3, value="")
            nc.border = border
            nc.fill = PatternFill("solid", start_color="FAFAFA") if idx % 2 == 0 else PatternFill(fill_type=None)

            lines_n = max(1, len(val)//80 + val.count("\n"))
            ws.row_dimensions[r].height = max(20, min(lines_n * 15, 400))

        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 75
        ws.column_dimensions["C"].width = 25
        ws.freeze_panes = "A3"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ══════════════════════════════════════════════════════════════════════════
# CHAT
# ══════════════════════════════════════════════════════════════════════════

STOP = {"the","and","for","with","this","that","from","have","will","shall",
        "must","should","would","their","which","where","when","what","into",
        "through","between","based","such","also","each","all","any","are",
        "was","were","has","not","but","more","than","its","our","your"}

def chat_answer(client, lines, full_text, history, question):
    q_words = [w for w in question.lower().split() if len(w) > 3 and w not in STOP]
    snippet = search_pdf(lines, q_words, ctx=20, max_s=4)
    ctx     = clean(full_text[:2000])
    if snippet:
        ctx += "\n\n--- RELEVANT SECTIONS ---\n" + snippet[:3000]

    messages = [{
        "role":    "system",
        "content": ("You are a procurement expert. Answer ONLY based on the tender "
                    "document below. Quote exact figures, dates, clause numbers.\n\n"
                    "TENDER:\n" + ctx)
    }]
    for m in history:
        messages.append({"role": "user" if m["role"] == "user" else "assistant",
                         "content": m["text"]})
    messages.append({"role": "user", "content": question})

    r = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=messages,
        max_tokens=900,
        temperature=0.1
    )
    return r.choices[0].message.content


# ══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════

for k, v in {
    "mode": None, "chat_history": [], "pdf_lines": None,
    "pdf_text": None, "pdf_name": None,
    "last_excel": None, "last_results": None,
    "history": [], "page": "main",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📄 Tender Analyzer")
    st.markdown("---")
    ca, cb = st.columns(2)
    with ca:
        if st.button("🏠 Home", use_container_width=True,
                     type="primary" if st.session_state.page == "main" else "secondary"):
            st.session_state.page = "main"; st.rerun()
    with cb:
        n = len(st.session_state.history)
        if st.button(f"🕓 History ({n})", use_container_width=True,
                     type="primary" if st.session_state.page == "history" else "secondary"):
            st.session_state.page = "history"; st.rerun()

    st.markdown("---")
    st.markdown("### Upload Tender PDF")
    pdf_file = st.file_uploader("Any government tender PDF", type=["pdf"])

    if pdf_file:
        if st.session_state.pdf_name != pdf_file.name:
            with st.spinner("Reading PDF…"):
                lines, full = read_pdf(pdf_file)
                st.session_state.pdf_lines    = lines
                st.session_state.pdf_text     = full
                st.session_state.pdf_name     = pdf_file.name
                st.session_state.chat_history = []
                st.session_state.mode         = None
                st.session_state.last_excel   = None
                st.session_state.last_results = None
        st.success(f"✅ {pdf_file.name}")
        st.caption(f"{len(st.session_state.pdf_lines):,} lines extracted")

    st.markdown("---")
    if st.session_state.mode and st.session_state.page == "main":
        if st.button("🔄 Change Mode", use_container_width=True):
            st.session_state.mode         = None
            st.session_state.chat_history = []
            st.session_state.last_excel   = None
            st.session_state.last_results = None
            st.rerun()

    st.markdown(" ")
    st.markdown(
        "<small style='color:#aaa'>Works with ANY government tender PDF.<br>"
        "No template needed — generates Excel automatically.</small>",
        unsafe_allow_html=True
    )


# ══════════════════════════════════════════════════════════════════════════
# HISTORY PAGE
# ══════════════════════════════════════════════════════════════════════════

if st.session_state.page == "history":
    st.markdown("<p class='page-title'>🕓 Tender History</p>", unsafe_allow_html=True)
    if not st.session_state.history:
        st.info("No tenders analyzed yet.")
        st.stop()
    for h in st.session_state.history:
        with st.expander(f"📄 {h['pdf_name']}  ·  {h['timestamp']}", expanded=False):
            c1, c2, _ = st.columns([2, 2, 4])
            with c1:
                if h.get("excel_bytes"):
                    st.download_button("⬇️ Download Excel",
                        data=h["excel_bytes"],
                        file_name=f"Tender_{h['pdf_name'].replace('.pdf','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{h['id']}")
            with c2:
                if st.button("🗑️ Remove", key=f"rm_{h['id']}"):
                    st.session_state.history = [x for x in st.session_state.history
                                                 if x["id"] != h["id"]]
                    st.rerun()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
# MAIN PAGE
# ══════════════════════════════════════════════════════════════════════════

st.markdown("<p class='page-title'>📄 Tender Document Analyzer</p>", unsafe_allow_html=True)
st.markdown(
    "<p class='page-sub'>Upload any government tender PDF — AI reads it and creates a structured Excel report</p>",
    unsafe_allow_html=True)

client = Groq(api_key=GROQ_API_KEY)

if not pdf_file:
    st.markdown(
        "<div class='info-box'>👈 <b>Upload any tender PDF</b> from the sidebar to get started.</div>",
        unsafe_allow_html=True)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown("""<div style='background:white;border-radius:16px;padding:1.5rem;
            border:2px solid #e8eaf6;text-align:center'>
            <div style='font-size:2.5rem'>💬</div>
            <div style='font-size:1.1rem;font-weight:700;margin:.5rem 0'>Chat Mode</div>
            <div style='font-size:.85rem;color:#555'>Ask any question about the tender</div>
            </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""<div style='background:white;border-radius:16px;padding:1.5rem;
            border:2px solid #e8eaf6;text-align:center'>
            <div style='font-size:2.5rem'>📊</div>
            <div style='font-size:1.1rem;font-weight:700;margin:.5rem 0'>Auto-Analyze Mode</div>
            <div style='font-size:.85rem;color:#555'>AI extracts 50+ fields and generates Excel</div>
            </div>""", unsafe_allow_html=True)
    st.stop()

lines = st.session_state.pdf_lines
text  = st.session_state.pdf_text

# Mode selection
if st.session_state.mode is None:
    st.markdown("### What would you like to do?")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown("""<div style='background:white;border-radius:16px;padding:1.5rem;
            border:2px solid #e8eaf6;text-align:center'>
            <div style='font-size:2.5rem'>💬</div>
            <div style='font-size:1.1rem;font-weight:700;margin:.5rem 0'>Chat with the Tender</div>
            <div style='font-size:.85rem;color:#555'>Ask anything — EMD, eligibility, deadlines, scope, payment</div>
            </div>""", unsafe_allow_html=True)
        st.markdown("")
        if st.button("💬 Start Chatting", use_container_width=True, type="primary"):
            st.session_state.mode = "chat"; st.rerun()
    with c2:
        st.markdown("""<div style='background:white;border-radius:16px;padding:1.5rem;
            border:2px solid #e8eaf6;text-align:center'>
            <div style='font-size:2.5rem'>📊</div>
            <div style='font-size:1.1rem;font-weight:700;margin:.5rem 0'>Auto-Analyze & Export Excel</div>
            <div style='font-size:.85rem;color:#555'>AI extracts 50+ fields across 6 sections, creates professional Excel report</div>
            </div>""", unsafe_allow_html=True)
        st.markdown("")
        if st.button("📊 Analyze & Generate Excel", use_container_width=True):
            st.session_state.mode = "analyze"; st.rerun()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
# CHAT MODE
# ══════════════════════════════════════════════════════════════════════════

if st.session_state.mode == "chat":
    st.markdown(f"### 💬 Chatting with: {pdf_file.name}")

    with st.expander("🔍 Quick questions", expanded=True):
        quick = [
            "Who issued this tender?",
            "What is the tender number?",
            "What is the EMD amount?",
            "What is the bid closing date?",
            "What are the pre-qualification criteria?",
            "What is the scope of work?",
            "What are the payment terms?",
            "What is the CPBG requirement?",
            "What is the delivery period?",
        ]
        clicked = None
        for row in [quick[i:i+3] for i in range(0, len(quick), 3)]:
            cols = st.columns(3)
            for j, q in enumerate(row):
                with cols[j]:
                    if st.button(q, key=f"q_{q[:12]}", use_container_width=True):
                        clicked = q

    st.markdown("---")
    for msg in st.session_state.chat_history:
        if msg["role"] == "user":
            st.markdown(
                f"<div class='chat-user-wrap'><div class='chat-user'>{msg['text']}</div></div>",
                unsafe_allow_html=True)
        else:
            st.markdown(
                f"<div class='chat-bot-wrap'><div class='bot-avatar'>🤖</div>"
                f"<div class='chat-bot'>{msg['text']}</div></div>",
                unsafe_allow_html=True)

    q = clicked or st.chat_input("Ask anything about this tender…")
    if q:
        st.session_state.chat_history.append({"role": "user", "text": q})
        with st.spinner("Thinking…"):
            try:
                ans = chat_answer(client, lines, text,
                                  st.session_state.chat_history[:-1], q)
            except Exception as e:
                ans = f"Error: {e}"
        st.session_state.chat_history.append({"role": "bot", "text": ans})
        st.rerun()

    if st.session_state.chat_history:
        if st.button("🗑️ Clear chat"):
            st.session_state.chat_history = []; st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# ANALYZE MODE
# ══════════════════════════════════════════════════════════════════════════

elif st.session_state.mode == "analyze":
    st.markdown(f"### 📊 Analyzing: {pdf_file.name}")

    with st.expander("📋 What will be extracted (6 sections, 50+ fields)", expanded=False):
        for sec in SECTIONS:
            st.markdown(f"**{sec['title']}** — {len(sec['fields'])} fields")
            for fname, _ in sec["fields"]:
                st.write(f"  • {fname}")

    st.markdown("")

    if st.button("🤖 Start Analysis & Generate Excel", type="primary", use_container_width=True):
        st.session_state.last_excel   = None
        st.session_state.last_results = None

        bar       = st.progress(0, text="Starting…")
        status_ph = st.empty()

        try:
            all_results = extract_all(client, lines, text, bar, status_ph)

            bar.progress(94, text="Generating Excel report…")
            status_ph.empty()

            xl_bytes = make_excel(all_results, st.session_state.pdf_name).getvalue()
            st.session_state.last_excel   = xl_bytes
            st.session_state.last_results = all_results

            entry = {
                "id":          len(st.session_state.history) + 1,
                "pdf_name":    st.session_state.pdf_name,
                "mode":        "analyze",
                "excel_bytes": xl_bytes,
                "timestamp":   datetime.now().strftime("%d %b %Y, %I:%M %p"),
                "exported":    True,
            }
            st.session_state.history = [h for h in st.session_state.history
                                         if h["pdf_name"] != st.session_state.pdf_name]
            st.session_state.history.insert(0, entry)
            bar.progress(100, text="Done! ✅")

        except Exception as e:
            bar.empty(); status_ph.empty()
            st.error(f"Error: {e}")
            import traceback; st.code(traceback.format_exc())
            st.stop()

    if st.session_state.last_results and st.session_state.last_excel:
        res     = st.session_state.last_results
        found   = sum(1 for d in res.values()
                      for v in d.values() if "not specified" not in v.lower())
        total   = sum(len(d) for d in res.values())
        missing = total - found

        st.markdown(f"""<div class='success-box'>
        ✅ <b>Analysis complete!</b> &nbsp;
        <b>{found}</b> fields extracted &nbsp;·&nbsp;
        <b>{missing}</b> not in this PDF &nbsp;·&nbsp;
        <b>{total}</b> total fields checked
        </div>""", unsafe_allow_html=True)

        if missing > 0:
            st.markdown(
                f"<div class='warn-box'>ℹ️ <b>{missing} fields</b> show "
                f"'Not specified in document' — this is normal. "
                f"Different tenders have different sections.</div>",
                unsafe_allow_html=True)

        # Results preview tabs
        tabs = st.tabs(["📊 Summary"] + [s["title"] for s in SECTIONS])

        with tabs[0]:
            key_fields = [
                ("Issuing Organization",         "overview"),
                ("Tender No / NIT No",           "overview"),
                ("GeM Bid Number",               "overview"),
                ("Bid Submission Deadline",      "overview"),
                ("EMD / Bid Security",           "overview"),
                ("Offer Validity",               "overview"),
                ("Delivery / Completion Period", "commercial"),
                ("Payment Terms",               "commercial"),
                ("Warranty Period",             "commercial"),
                ("CPBG / Performance Security", "commercial"),
                ("Package Summary (Nutshell)",  "scope"),
                ("Area / Location of Work",     "scope"),
                ("Evaluation Method",           "bec"),
                ("Rejection Criteria",          "bec"),
            ]
            rows = ""
            for fname, sid in key_fields:
                val = res.get(sid, {}).get(fname, "Not specified in document")
                nf  = "not specified" in val.lower()
                vs  = val.replace("\n", "<br>")[:300]
                cls = "missing" if nf else "found"
                rows += f"<tr><td style='font-weight:600;width:230px'>{fname}</td><td class='{cls}'>{vs}</td></tr>"
            st.markdown(f"<table class='preview-table'>{rows}</table>", unsafe_allow_html=True)

        for i, section in enumerate(SECTIONS):
            with tabs[i + 1]:
                sid  = section["id"]
                rows = ""
                for fname, _ in section["fields"]:
                    val = res.get(sid, {}).get(fname, "Not specified in document")
                    nf  = "not specified" in val.lower()
                    vs  = val.replace("\n", "<br>").replace("&", "&amp;")[:500]
                    cls = "missing" if nf else "found"
                    rows += f"<tr><td style='font-weight:600;width:230px'>{fname}</td><td class='{cls}'>{vs}</td></tr>"
                st.markdown(f"<table class='preview-table'>{rows}</table>", unsafe_allow_html=True)

        st.markdown("")
        st.download_button(
            label="⬇️ Download Excel Report",
            data=st.session_state.last_excel,
            file_name=f"Tender_{st.session_state.pdf_name.replace('.pdf','')}_{datetime.now().strftime('%d%b%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )