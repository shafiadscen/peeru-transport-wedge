from datetime import date, datetime, timezone, timedelta
import io
import dropbox
import openpyxl
import streamlit as st


# Define Oman time zone (UTC +4)
oman_tz = timezone(timedelta(hours=4))


# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="Peeru Transport",
    page_icon="🚚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Custom dark-theme styling matching your desktop widget
st.markdown(
    """
    <style>
    .stApp { background-color: #12161E; color: #E2E8F0; }
    .card {
        background-color: #1B2230;
        border: 1px solid #334155;
        padding: 14px;
        border-radius: 10px;
        margin-bottom: 10px;
    }
    .metric-title { font-size: 13px; color: #94A3B8; font-weight: bold; }
    .metric-value { font-size: 22px; font-weight: bold; color: #E2E8F0; }
    .metric-sub { font-size: 11px; color: #94A3B8; }
    .row-card {
        background-color: #1B2230;
        border-left: 4px solid #334155;
        padding: 10px 12px;
        border-radius: 6px;
        margin-bottom: 8px;
        font-size: 13px;
    }
    .due-row { border-left-color: #F59E0B; }
    .ok-row { border-left-color: #10B981; }
    .out-row { border-left-color: #EF4444; }
    .in-row { border-left-color: #10B981; }
    </style>
""",
    unsafe_allow_html=True,
)

# ================= SIMPLE PIN SECURITY =================
SECRET_PIN = "2815"  # Change this to your preferred PIN

if "authenticated" not in st.session_state:
  st.session_state.authenticated = False

if not st.session_state.authenticated:
  st.markdown(
      "<h2 style='text-align: center; color: #38BDF8;'>🔒 Peeru Transport"
      " Login</h2>",
      unsafe_allow_html=True,
  )
  entered_pin = st.text_input("Enter Secret PIN", type="password")
  if st.button("Login", use_container_width=True):
    if entered_pin == SECRET_PIN:
      st.session_state.authenticated = True
      st.rerun()
    else:
      st.error("Incorrect PIN. Access Denied.")
  st.stop()

# ================= CONFIG & CREDENTIALS =================
DROPBOX_APP_KEY = "oy92bbzbbcwsrs9"
DROPBOX_APP_SECRET = "9m0thk565j2sd4x"
DROPBOX_REFRESH_TOKEN = (
    "qbnr91hHnKoAAAAAAAAAASnjBV9r7NYagJURxKQd18EfcjFwSJo73YFl-yUay3sp"
)
DROPBOX_NS_PATH = "ns:3077906689//DAILY CASH DETAIL/DAILY CASH-CREDIT.xlsx"

TABS = [
    ("cash", "CASH", "Cash-Credit", "cash"),
    ("nadeem", "NADEEM TR", "Nadeem TR", "tr"),
    ("mukhtar", "MUKHTAR TR", "Mukhtar TR", "tr"),
    ("safdar", "SAFDAR TR", "Safdar TR", "tr"),
]


def num(v):
  return (
      v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
  )


def blank(v):
  return v is None or (isinstance(v, str) and not v.strip())


def s(v):
  return "" if blank(v) else str(v).strip()


def fdate(v):
  if isinstance(v, (datetime, date)):
    return v.strftime("%d-%b-%Y")
  return s(v)[:11]


def f3(v, plus=False):
  if num(v) is None:
    return "-"
  return f"{v:+,.3f}" if plus else f"{v:,.3f}"


def is_due(e):
  return bool(e.get("bal", 0) and e.get("bal", 0) > 0.0005)


# ================= FULL DATA PARSER =================
@st.cache_data(ttl=30)
def load_cloud_data():
  dbx = dropbox.Dropbox(
      app_key=DROPBOX_APP_KEY,
      app_secret=DROPBOX_APP_SECRET,
      oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
  )
  
  metadata = dbx.files_get_metadata(DROPBOX_NS_PATH)
  last_modified_utc = metadata.server_modified
  last_modified_oman = last_modified_utc.replace(tzinfo=timezone.utc).astimezone(oman_tz)
  file_updated_str = last_modified_oman.strftime("%d-%b-%Y %H:%M")

  _, res = dbx.files_download(DROPBOX_NS_PATH)
  wb = openpyxl.load_workbook(
      io.BytesIO(res.content), data_only=True, read_only=True
  )

  out = {"file_updated": file_updated_str}
  try:
    for key, _, sheet, kind in TABS:
      ws = None
      for n in wb.sheetnames:
        if n.strip().lower() == sheet.lower():
          ws = wb[n]
          break
      if not ws:
        continue

      if kind == "cash":
        summary, entries = None, []
        for i, row in enumerate(
            ws.iter_rows(
                min_row=1, max_row=50000, max_col=9, values_only=True
            ),
            start=1,
        ):
          if i == 2:
            summary = num(row[8])
          if i < 3 or blank(row[1]):
            continue
          
          # Normalize date safely for comparison
          raw_date = row[4]
          entries.append({
              "date": raw_date,
              "party": s(row[1]),
              "typ": s(row[2]),
              "det": s(row[3]),
              "amt": num(row[0]) or 0,
              "bal": num(row[5]) or 0,
          })
        balance = summary if summary is not None else (entries[-1]["bal"] if entries else 0)
        out[key] = {"balance": balance or 0, "entries": entries[::-1]}

      else:  # Transporter sheets
        balance = ctns = None
        entries = []
        for i, row in enumerate(
            ws.iter_rows(
                min_row=1, max_row=3000, max_col=12, values_only=True
            ),
            start=1,
        ):
          if i == 2:
            balance = num(row[11])
          elif i == 3:
            ctns = num(row[11])
          if i < 5 or (blank(row[4]) and num(row[8]) is None):
            continue
          entries.append({
              "payon": row[1],
              "rcpt": s(row[2]),
              "desc": s(row[3]),
              "cont": s(row[4]),
              "load": row[5],
              "unl": row[6],
              "ctns": num(row[7]) or 0,
              "amt": num(row[8]) or 0,
              "disc": num(row[9]) or 0,
              "paid": num(row[10]) or 0,
              "bal": num(row[11]) or 0,
          })
        unpaid = [e for e in entries if is_due(e)]
        if balance is None:
          balance = sum(e["bal"] for e in unpaid)
        out[key] = {
            "balance": balance or 0,
            "ctns": int(ctns or 0),
            "unpaid": len(unpaid),
            "entries": entries[::-1],
        }
  finally:
    wb.close()
  return out

# ================= UI HEADER =================
st.markdown(
    "<h2 style='text-align: center; color: #38BDF8; margin-bottom:"
    " 0px;'>PEERU TR</h2>",
    unsafe_allow_html=True,
)

if st.button("⟳ Fetch Latest from Cloud", use_container_width=True):
  st.cache_data.clear()
  st.rerun()

with st.spinner("Syncing with Dropbox..."):
  data = load_cloud_data()

file_saved_time = data.get("file_updated", "Unknown")

st.markdown(
    f"<p style='text-align: center; color: #94A3B8; font-size: 11px; margin-bottom: 0px;'>Full Mobile Report • Viewed: {datetime.now(oman_tz):%d-%b-%Y %H:%M}</p>",
    unsafe_allow_html=True,
)
st.markdown(
    f"<p style='text-align: center; color: #10B981; font-size: 11px;'>📁 Staff File Saved: {file_saved_time}</p>",
    unsafe_allow_html=True,
)

# ================= PRO EXECUTIVE COMMAND CENTER =================
st.markdown("---")
# Date selector defaulting to today (Oman time)
selected_eval_date = st.date_input(
    "📅 Select Date for Performance Audit",
    value=datetime.now(oman_tz).date()
)

eval_collections = 0.0
eval_expenses = 0.0
containers_eval = 0
cartons_eval = 0

# Scan cash entries for selected date
cash_info = data.get("cash", {})
for e in cash_info.get("entries", []):
  e_date = e.get("date")
  if isinstance(e_date, datetime):
    e_date = e_date.date()
  if e_date == selected_eval_date:
    amt = e.get("amt", 0)
    if amt > 0:
      eval_collections += amt
    else:
      eval_expenses += abs(amt)

# Scan transporter entries for selected date containers and cartons (based on unload date)
for t_key in ["nadeem", "mukhtar", "safdar"]:
  t_info = data.get(t_key, {})
  for e in t_info.get("entries", []):
    unl_date = e.get("unl")
    if isinstance(unl_date, datetime):
      unl_date = unl_date.date()
    if unl_date == selected_eval_date:
      containers_eval += 1
      cartons_eval += int(e.get("ctns", 0))

net_flow = eval_collections - eval_expenses
net_flow_color = "#10B981" if net_flow >= 0 else "#EF4444"
formatted_eval_date = selected_eval_date.strftime("%d-%b-%Y")

st.markdown(
    f"""
    <style>
    .exec-card {{
        background: linear-gradient(135deg, #1B2230 0%, #0F172A 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    }}
    .exec-header {{
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #38BDF8;
        font-weight: 700;
        margin-bottom: 10px;
        border-bottom: 1px solid #1E293B;
        padding-bottom: 6px;
    }}
    .exec-grid {{
        display: flex;
        justify-content: space-between;
        gap: 6px;
    }}
    .exec-metric {{
        flex: 1;
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid #1E293B;
        padding: 8px 4px;
        border-radius: 8px;
        text-align: center;
    }}
    .exec-label {{
        font-size: 9px;
        color: #94A3B8;
        margin-bottom: 4px;
        text-transform: uppercase;
    }}
    .exec-val {{
        font-size: 13px;
        font-weight: bold;
        color: #E2E8F0;
    }}
    </style>

    <div class="exec-card">
        <div class="exec-header">⚡ Executive Command Center • {formatted_eval_date}</div>
        <div class="exec-grid">
            <div class="exec-metric">
                <div class="exec-label">Inflow</div>
                <div class="exec-val" style="color: #10B981;">+{eval_collections:,.3f}</div>
            </div>
            <div class="exec-metric">
                <div class="exec-label">Outflow</div>
                <div class="exec-val" style="color: #EF4444;">-{eval_expenses:,.3f}</div>
            </div>
            <div class="exec-metric">
                <div class="exec-label">Containers</div>
                <div class="exec-val" style="color: #38BDF8;">{containers_eval} <span style="font-size:10px; color:#94A3B8;">({cartons_eval:,} ctns)</span></div>
            </div>
            <div class="exec-metric">
                <div class="exec-label">Net Flow</div>
                <div class="exec-val" style="color: {net_flow_color};">{net_flow:+,.3f}</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

================= Resend Email setup =================
import resend

resend.api_key = st.secrets.get("RESEND_API_KEY", "")

if st.button("📧 Email Executive Report for Selected Date", use_container_width=True):
  if not resend.api_key:
    st.error("Resend API key is missing from Streamlit secrets.")
  else:
    try:
      report_date_str = selected_eval_date.strftime("%d-%b-%Y")
      
      inflow_lines = []
      for e in cash_info.get("entries", []):
        e_date = e.get("date")
        if isinstance(e_date, datetime): e_date = e_date.date()
        if e_date == selected_eval_date and e.get("amt", 0) > 0:
          inflow_lines.append(f"• {e.get('party')}: {e.get('amt'):,.3f} OMR ({e.get('det')})")
      
      outflow_lines = []
      for e in cash_info.get("entries", []):
        e_date = e.get("date")
        if isinstance(e_date, datetime): e_date = e_date.date()
        if e_date == selected_eval_date and e.get("amt", 0) < 0:
          outflow_lines.append(f"• {e.get('party')}: {abs(e.get('amt')):,.3f} OMR ({e.get('det')})")

      container_lines = []
      for t_key in ["nadeem", "mukhtar", "safdar"]:
        for e in data.get(t_key, {}).get("entries", []):
          unl_val = e.get("unl")
          if isinstance(unl_val, datetime): unl_val = unl_val.date()
          if unl_val == selected_eval_date:
            container_lines.append(f"• {t_key.upper()} - Container {e.get('cont')}: {e.get('ctns')} cartons")

      email_html = f"""
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 10px;">
          <h2 style="color: #0284c7; text-align: center;">⚡ PEERU TRANSPORT COMMAND CENTER</h2>
          <p style="text-align: center; color: #666; font-size: 14px;">Executive Daily Report for <b>{report_date_str}</b></p>
          <hr style="border: none; border-top: 1px solid #eee;" />
          
          <h3>1. Inflow (+ Total Collections)</h3>
          <ul>{''.join(inflow_lines) if inflow_lines else '<li>No Inflows recorded</li>'}</ul>
          <p><b>Total Inflow:</b> +{eval_collections:,.3f} OMR</p>

          <h3>2. Outflow (- Expenses & Payments)</h3>
          <ul>{''.join(outflow_lines) if outflow_lines else '<li>No Outflows recorded</li>'}</ul>
          <p><b>Total Outflow:</b> -{eval_expenses:,.3f} OMR</p>

          <h3>3. Containers & Cartons Unloaded</h3>
          <ul>{''.join(container_lines) if container_lines else '<li>No Containers unloaded</li>'}</ul>
          <p><b>Total Containers:</b> {containers_eval} ({cartons_eval:,} ctns)</p>

          <hr style="border: none; border-top: 1px solid #eee;" />
          <h3>4. Net Cash Flow Summary</h3>
          <p style="font-size: 16px;"><b>Net Flow:</b> <span style="color: {'#10b981' if net_flow >= 0 else '#ef4444'};">{net_flow:+,.3f} OMR</span></p>
      </div>
      """

      params = {
          "from": "Peeru Transport <onboarding@resend.dev>",
          "to": [st.secrets.get("TRANSPORT_EMAIL_RECIPIENT", "")],
          "subject": f"Peeru Transport Executive Report — {report_date_str}",
          "html": email_html,
      }
      
      resend.Emails.send(params)
      st.success(f"🚀 Executive Report emailed successfully for {report_date_str}!")
    except Exception as ex:
      st.error(f"Failed to send email: {ex}")

# ================= SUMMARY CARDS =================
for key, label, _, kind in TABS:
  info = data.get(key, {"balance": 0})
  bal = info.get("balance", 0)

  if kind == "cash":
    subtext = "Main Cash Balance"
  else:
    subtext = (
        f"{info.get('unpaid', 0)} Unpaid Bills  •  {info.get('ctns', 0):,} Ctns"
    )

  color_style = "color: #10B981;" if bal >= 0 else "color: #EF4444;"

  st.markdown(
      f"""
        <div class="card">
            <div class="metric-title">{label}</div>
            <div class="metric-value" style="{color_style}">{bal:,.3f} <span style="font-size:13px; color:#94A3B8;">OMR</span></div>
            <div class="metric-sub">{subtext}</div>
        </div>
    """,
      unsafe_allow_html=True,
  )

st.markdown("---")

# ================= DETAILED LEDGER & SEARCH VIEW =================
selected_tab = st.selectbox(
    "Select Ledger to Inspect", [t[1] for t in TABS], index=0
)
key_map = {t[1]: t[0] for t in TABS}
active_key = key_map[selected_tab]
kind = [t[3] for t in TABS if t[0] == active_key][0]
active_info = data.get(active_key, {})
entries = active_info.get("entries", [])

# Filters for Transporter sheets
filter_mode = "All entries"
if kind == "tr":
  filter_mode = st.radio(
      "Filter Transporter View",
      ["All entries", "Unpaid only"],
      horizontal=True,
      label_visibility="collapsed",
  )
  if filter_mode == "Unpaid only":
    entries = [e for e in entries if is_due(e)]

# Search Bar
search_query = st.text_input(
    "🔍 Search entries...", placeholder="Type name, container#, receipt#..."
).strip().lower()

if search_query:
  filtered_entries = []
  for e in entries:
    if kind == "cash":
      match_text = f"{e['party']} {e['typ']} {e['det']}".lower()
    else:
      match_text = f"{e['cont']} {e['desc']} {e['rcpt']}".lower()

    if search_query in match_text:
      filtered_entries.append(e)
  entries = filtered_entries

st.subheader(f"{selected_tab} ({len(entries)} items found)")

if not entries:
  st.info("No records matching your search.")
else:
  for item in entries:
    if kind == "cash":
      tag = "out-row" if item["amt"] < 0 else "in-row"
      amt_col = "#EF4444" if item["amt"] < 0 else "#10B981"
      st.markdown(
          f"""
            <div class="row-card {tag}">
                <b>{fdate(item['date'])}</b> &nbsp;|&nbsp; <span style='color:#38BDF8;'>{item['party']}</span><br>
                <span style='color:#94A3B8;'>Type:</span> {item['typ']} &nbsp;|&nbsp; <span style='color:#94A3B8;'>Details:</span> {item['det']}<br>
                <b>Amount:</b> <span style='color:{amt_col};'>{f3(item['amt'], True)}</span> &nbsp;|&nbsp; <b>Balance:</b> {f3(item['bal'])}
            </div>
            """,
          unsafe_allow_html=True,
      )
    else:  # Transporter card
      due = is_due(item)
      tag = "due-row" if due else "ok-row"
      bal_col = "#F59E0B" if due else "#10B981"
      st.markdown(
          f"""
            <div class="row-card {tag}">
                <b>Container: {item['cont']}</b> &nbsp;|&nbsp; <span style='color:#94A3B8;'>Unloaded:</span> {fdate(item['unl'])}<br>
                <span style='color:#94A3B8;'>Desc:</span> {item['desc']} &nbsp;|&nbsp; <span style='color:#94A3B8;'>Receipt#:</span> {item['rcpt']}<br>
                <b>Cartons:</b> {int(item['ctns'])} &nbsp;|&nbsp; <b>Amount:</b> {f3(item['amt'])} &nbsp;|&nbsp; <b>Paid:</b> {f3(item['paid'])}<br>
                <b>Balance:</b> <span style='color:{bal_col}; font-weight:bold;'>{f3(item['bal'])}</span> &nbsp;|&nbsp; <span style='color:#94A3B8;'>Paid On:</span> {fdate(item['payon'])}
            </div>
            """,
          unsafe_allow_html=True,
      )