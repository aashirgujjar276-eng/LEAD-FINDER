"""
Lead Finder Web App — Velnex AI
=================================
Streamlit app: pick country + city + business type, scrape via Apify's
Google Maps Scraper, filter/prioritize, and download as Excel.

RUN LOCALLY
-----------
pip install streamlit apify-client openpyxl pandas
streamlit run app.py

DEPLOY
------
Push to GitHub, deploy on Streamlit Community Cloud (same as your
velnexai repo). Add your Apify token as a Streamlit "Secret" instead
of pasting it in the UI every time — see bottom of this file.
"""

import re
import io
import time
import datetime
import streamlit as st
import pandas as pd
from apify_client import ApifyClient
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Velnex Lead Finder", page_icon="🎯", layout="wide")

CUSTOM_CSS = """
<style>
:root {
    --navy-dark: #0a2540;
    --blue-deep: #0d3d75;
    --blue-mid: #1a63b8;
    --blue-bright: #2f8fe0;
    --blue-light: #bfe0fb;
}

/* App background gradient */
.stApp {
    background: linear-gradient(160deg, var(--navy-dark) 0%, var(--blue-mid) 55%, var(--blue-light) 100%);
}

/* Decorative network-node motif, bottom right, like the reference image */
.stApp::after {
    content: "";
    position: fixed;
    bottom: 24px;
    right: 24px;
    width: 130px;
    height: 130px;
    pointer-events: none;
    z-index: 0;
    background-image:
        radial-gradient(circle, rgba(255,255,255,0.85) 2px, transparent 2.5px),
        radial-gradient(circle, rgba(255,255,255,0.85) 2px, transparent 2.5px),
        radial-gradient(circle, rgba(255,255,255,0.85) 2px, transparent 2.5px),
        radial-gradient(circle, rgba(255,255,255,0.85) 2px, transparent 2.5px),
        radial-gradient(circle, rgba(255,255,255,0.85) 2px, transparent 2.5px);
    background-position: 10px 10px, 60px 5px, 100px 40px, 40px 70px, 90px 100px;
    background-repeat: no-repeat;
}

/* Title + captions */
h1, h2, h3 {
    color: #ffffff !important;
}
.stCaption, [data-testid="stCaptionContainer"] {
    color: var(--blue-light) !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(200deg, var(--navy-dark) 0%, var(--blue-deep) 100%);
}
[data-testid="stSidebar"] * {
    color: #eaf4ff !important;
}

/* Cards / containers for inputs */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
.stSelectbox > div > div {
    background-color: rgba(255,255,255,0.92) !important;
    border-radius: 8px !important;
    border: 1px solid var(--blue-mid) !important;
}

/* Primary button */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(90deg, var(--blue-deep) 0%, var(--blue-bright) 100%);
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: 600;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
    background: linear-gradient(90deg, var(--blue-mid) 0%, var(--blue-bright) 100%);
    border: none;
}

/* Dataframe container */
[data-testid="stDataFrame"] {
    background: rgba(255,255,255,0.95);
    border-radius: 10px;
    padding: 4px;
}

[data-testid="stAlert"] {
    border-radius: 8px;
}

/* Search button: red, distinct from download button */
.st-key-search_btn_wrap .stButton > button[kind="primary"] {
    background: linear-gradient(90deg, #b3261e 0%, #e03e35 100%) !important;
}
.st-key-search_btn_wrap .stButton > button[kind="primary"]:hover {
    background: linear-gradient(90deg, #931d17 0%, #c9342b 100%) !important;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

CHAIN_KEYWORDS_DEFAULT = [
    "aspen dental", "heartland dental", "smile direct", "gentle dental",
    "pacific dental", "western dental", "great expressions", "coast dental",
    "midwest dental", "comfort dental", "supercuts", "great clips",
    "massage envy", "fantastic sams",
]

COUNTRIES = ["United States", "United Kingdom", "Canada", "Pakistan", "Australia", "Other"]

BUSINESS_PRESETS = [
    "Dental clinic", "HVAC contractor", "Law firm", "Hair salon", "Spa",
    "Plumber", "Roofing company", "Physical therapy clinic", "Custom...",
]

COLUMNS = [
    "City", "Business Name", "Phone", "Website", "Address",
    "Rating", "Review Count", "Google Maps URL", "Category", "Status", "Notes",
]

# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def clean_phone(phone: str) -> str:
    if not phone:
        return ""
    digits = re.sub(r"\D", "", phone)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[0:3]}) {digits[3:6]}-{digits[6:]}"
    return phone


def is_chain(name: str, chain_keywords: list) -> bool:
    name_l = name.lower()
    return any(chain in name_l for chain in chain_keywords)


def scrape_city(client: ApifyClient, search_term: str, location: str, max_results: int, progress_cb=None):
    run_input = {
        "searchStringsArray": [search_term],
        "locationQuery": location,
        "maxCrawledPlacesPerSearch": max_results,
        "language": "en",
        "skipClosedPlaces": True,
    }
    run = client.actor("compass/crawler-google-places").call(run_input=run_input)
    items = list(client.dataset(run["defaultDatasetId"]).iterate_items())
    return items


def process_items(items: list, city: str, min_reviews: int, min_rating: float, chain_keywords: list) -> list:
    leads = []
    for item in items:
        name = item.get("title", "").strip()
        if not name or is_chain(name, chain_keywords):
            continue
        reviews = item.get("reviewsCount") or 0
        rating = item.get("totalScore") or 0
        if reviews < min_reviews or rating < min_rating:
            continue
        leads.append({
            "City": city,
            "Business Name": name,
            "Phone": clean_phone(item.get("phone", "")),
            "Website": item.get("website", ""),
            "Address": item.get("address", ""),
            "Rating": round(rating, 1),
            "Review Count": reviews,
            "Google Maps URL": item.get("url", ""),
            "Category": item.get("categoryName", ""),
            "Status": "Not Contacted",
            "Notes": "",
        })
    return leads


def build_excel(leads: list) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="Arial", size=10)
    status_options = '"Not Contacted,Emailed,Called,Demo Booked,Demo Done,Won,Lost,Not Interested"'

    ws = wb.create_sheet("All Leads")
    for col_idx, header in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    leads_sorted = sorted(leads, key=lambda r: (r["Review Count"], r["Rating"]), reverse=True)
    for r_idx, row in enumerate(leads_sorted, start=2):
        for c_idx, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=row[col])
            cell.font = body_font

    last_row = len(leads_sorted) + 1
    widths = [16, 30, 16, 28, 34, 9, 12, 34, 20, 16, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    if last_row > 1:
        dv = DataValidation(type="list", formula1=status_options, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"J2:J{last_row}")

        ws.conditional_formatting.add(
            f"F2:F{last_row}",
            ColorScaleRule(start_type="min", start_color="F8696B",
                            mid_type="percentile", mid_value=50, mid_color="FFEB84",
                            end_type="max", end_color="63BE7B"),
        )
        ws.conditional_formatting.add(
            f"J2:J{last_row}",
            CellIsRule(operator="equal", formula=['"Won"'],
                       fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")),
        )
        ws.conditional_formatting.add(
            f"J2:J{last_row}",
            CellIsRule(operator="equal", formula=['"Lost"'],
                       fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")),
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────

st.title("🎯 Velnex Lead Finder")
st.caption("Pick a location and business type, get a filtered, ready-to-call lead sheet.")

with st.sidebar:
    st.header("Apify Connection")
    # Prefer Streamlit secrets if set, otherwise let user paste one
    try:
        default_token = st.secrets["APIFY_API_TOKEN"]
    except Exception:
        default_token = ""
    api_token = st.text_input("Apify API Token", value=default_token, type="password",
                               help="console.apify.com → Settings → Integrations")

    st.header("Filters")
    min_reviews = st.slider("Minimum reviews", 0, 200, 25)
    min_rating = st.slider("Minimum rating", 0.0, 5.0, 3.8, 0.1)
    exclude_chains = st.checkbox("Exclude known chains", value=True)
    max_results = st.slider("Max results per city", 10, 120, 60)

col1, col2, col3 = st.columns(3)
with col1:
    country = st.selectbox("Country", COUNTRIES)
with col2:
    cities_input = st.text_input("City / cities (comma-separated)", placeholder="Austin TX, Dallas TX")
with col3:
    business_preset = st.selectbox("Business type", BUSINESS_PRESETS)
    if business_preset == "Custom...":
        business_type = st.text_input("Enter business type", placeholder="e.g. orthodontist")
    else:
        business_type = business_preset

with st.container(key="search_btn_wrap"):
    run_button = st.button("🔍 Find Leads", type="primary", use_container_width=True)

if "leads" not in st.session_state:
    st.session_state.leads = []

if run_button:
    if not api_token:
        st.error("Enter your Apify API token in the sidebar first.")
    elif not cities_input.strip():
        st.error("Enter at least one city.")
    elif not business_type.strip():
        st.error("Enter or select a business type.")
    else:
        cities = [c.strip() for c in cities_input.split(",") if c.strip()]
        # Append country to each city query for accuracy, unless already implied
        locations = [f"{c}, {country}" if country != "Other" else c for c in cities]

        client = ApifyClient(api_token)
        all_leads = []
        progress = st.progress(0, text="Starting...")
        status = st.empty()

        for i, (city, location) in enumerate(zip(cities, locations)):
            status.text(f"Scraping {city}...")
            try:
                items = scrape_city(client, business_type, location, max_results)
                chain_list = CHAIN_KEYWORDS_DEFAULT if exclude_chains else []
                leads = process_items(items, city, min_reviews, min_rating, chain_list)
                all_leads.extend(leads)
                status.text(f"{city}: {len(leads)} qualified leads found")
            except Exception as e:
                st.warning(f"Error scraping {city}: {e}")
            progress.progress((i + 1) / len(cities))

        progress.empty()
        status.empty()
        st.session_state.leads = all_leads
        st.success(f"Done — {len(all_leads)} qualified leads across {len(cities)} location(s).")

if st.session_state.leads:
    df = pd.DataFrame(st.session_state.leads).sort_values(
        by=["Review Count", "Rating"], ascending=False
    )
    st.dataframe(df, use_container_width=True, hide_index=True)

    excel_bytes = build_excel(st.session_state.leads)
    filename = f"leads_{business_type.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "⬇️ Download Excel CRM",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
