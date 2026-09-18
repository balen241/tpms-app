import streamlit as st
import pandas as pd
import json
import os
import uuid
from datetime import datetime, date, time, timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# --- NEPAL TIMEZONE HELPERS ---
# Streamlit Cloud servers run on UTC, not Nepal time, so every "now" used for
# clocking in/out, shift detection, and default dates must go through these.
NEPAL_TZ = ZoneInfo("Asia/Kathmandu")

def nepal_now():
    """Current date/time in Nepal (naive, so it stores/compares cleanly as plain text)."""
    return datetime.now(NEPAL_TZ).replace(tzinfo=None)

def nepal_today():
    return nepal_now().date()

def determine_shift(dt):
    """Given a Nepal-local datetime, return the correct shift label based on actual clock time.
    Morning: 06:00–14:00 | Evening: 14:00–22:00 | Night: 22:00–06:00 (wraps past midnight)."""
    t = dt.time()
    if time(6, 0) <= t < time(14, 0):
        return "Morning Shift (06:00 - 14:00)"
    elif time(14, 0) <= t < time(22, 0):
        return "Evening Shift (14:00 - 22:00)"
    else:
        # Covers 22:00–23:59 and 00:00–05:59
        return "Night Shift (22:00 - 06:00)"

st.set_page_config(page_title="TPMS", page_icon="⏰", layout="wide")

st.markdown("""
<style>
    /* Overall spacing */
    .block-container { padding-top: 2rem; padding-bottom: 3rem; }

    /* Page title */
    h1 { font-weight: 700; letter-spacing: -0.5px; }

    /* Section headers */
    h3 { margin-top: 0.5rem; color: #1a1a2e; }
    h4 { color: #333; font-weight: 600; }

    /* Metric cards */
    div[data-testid="stMetric"] {
        background: #f8f9fb;
        border: 1px solid #e6e8ec;
        border-radius: 12px;
        padding: 1rem 1.25rem;
    }
    div[data-testid="stMetricLabel"] { font-weight: 500; color: #6b7280; }
    div[data-testid="stMetricValue"] { font-weight: 700; color: #1a1a2e; }

    /* Buttons */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        padding: 0.5rem 1rem;
    }

    /* Dataframes */
    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #e6e8ec;
    }

    /* Dividers */
    hr { margin: 1.25rem 0; }

    /* Alerts (success/info/warning/error) */
    div[data-testid="stAlert"] { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# --- GOOGLE SHEETS CONNECTION SETUP ---
@st.cache_resource
def init_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
    except Exception:
        json_files = [f for f in os.listdir('.') if f.endswith('.json') and 'payroll-and-management' in f]
        with open(json_files[0]) as f:
            creds_dict = json.load(f)

    # Normalize the private key so it's valid PEM regardless of how it was
    # stored in secrets.toml (fixes "Could not deserialize key data" errors
    # caused by escaped \n sequences not being converted to real newlines).
    pk = creds_dict.get("private_key", "")
    if "\\n" in pk and "\n" not in pk.replace("\\n", ""):
        pk = pk.replace("\\n", "\n")
    creds_dict["private_key"] = pk.strip() + "\n"

    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open("Payroll System Database")
    return sheet

gsheet = init_google_sheet()

# --- VERIFY AND INITIALIZE SHEET STRUCTURE (CACHED TO PREVENT 429 ERRORS) ---
@st.cache_resource
def verify_sheet_structure():
    required_tabs = {
        "employees": ["emp_id", "name", "hourly_rate", "pin", "shift_name"],
        "time_punches": ["punch_id", "emp_id", "shift_name", "clock_in", "clock_out", "total_hours", "cash_in", "cash_out", "bonus", "approval_status", "manager_message", "message_acknowledged"],
        "expenses": ["expense_id", "expense_date", "category", "vendor", "amount", "notes"],
        "settings": ["key", "value"],
        "shift_transactions": ["txn_id", "punch_id", "emp_id", "amount", "timestamp"],
        "active_sessions": ["session_token", "emp_id", "last_activity"]
    }
    
    try:
        existing_ws_map = {ws.title.lower(): ws for ws in gsheet.worksheets()}
    except Exception:
        existing_ws_map = {}
    
    for tab, headers in required_tabs.items():
        if tab.lower() in existing_ws_map:
            ws = existing_ws_map[tab.lower()]
            if not ws.row_values(1):
                ws.append_row(headers)
        else:
            try:
                ws = gsheet.add_worksheet(title=tab, rows="1000", cols="20")
                ws.append_row(headers)
            except Exception:
                ws = gsheet.worksheet(tab)
                if not ws.row_values(1):
                    ws.append_row(headers)

verify_sheet_structure()

# --- HELPER FUNCTIONS WITH CACHING ---
@st.cache_data(ttl=5)
def get_as_df(worksheet_name):
    try:
        ws = gsheet.worksheet(worksheet_name)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()

def safe_float(val):
    try:
        if val is None:
            return 0.0
        val_str = str(val).strip().replace('$', '').replace(',', '')
        if val_str == '' or val_str.lower() == 'nan':
            return 0.0
        return float(val_str)
    except:
        return 0.0

def execute_query(worksheet_name, action, data_row=None, row_id=None, update_dict=None):
    ws = gsheet.worksheet(worksheet_name)
    if action == "insert":
        ws.append_row(data_row)
    elif action == "update" and row_id is not None:
        cell = ws.find(str(row_id))
        if cell:
            row_num = cell.row
            headers = ws.row_values(1)
            for k, v in update_dict.items():
                if k in headers:
                    col_num = headers.index(k) + 1
                    ws.update_cell(row_num, col_num, v)
    st.cache_data.clear()

def delete_matching_rows(worksheet_name, column_name, value):
    """Delete every row in worksheet_name where column_name == value.
    Used to wipe temporary per-shift data (e.g. shift_transactions) once it's no longer needed."""
    ws = gsheet.worksheet(worksheet_name)
    all_values = ws.get_all_values()
    if not all_values:
        return
    headers = all_values[0]
    if column_name not in headers:
        return
    col_idx = headers.index(column_name)
    rows_to_delete = [
        i for i, row in enumerate(all_values[1:], start=2)
        if len(row) > col_idx and str(row[col_idx]) == str(value)
    ]
    for row_num in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_num)
    st.cache_data.clear()

def get_setting(key, default=""):
    try:
        ws = gsheet.worksheet("settings")
        records = ws.get_all_records()
        for r in records:
            if str(r.get("key", "")) == key:
                return str(r.get("value", ""))
    except Exception:
        pass
    return default

def set_setting(key, value):
    ws = gsheet.worksheet("settings")
    cell = ws.find(key)
    if cell:
        ws.update_cell(cell.row, 2, value)
    else:
        ws.append_row([key, value])
    st.cache_data.clear()

# --- PERSISTENT LOGIN SESSIONS (survive page refresh, expire after 30 min idle) ---
SESSION_TIMEOUT_MINUTES = 30

def create_session(emp_id):
    """Start a new session for this employee and return its token."""
    token = str(uuid.uuid4())
    ws = gsheet.worksheet("active_sessions")
    ws.append_row([token, emp_id, nepal_now().strftime("%Y-%m-%d %H:%M:%S")])
    st.cache_data.clear()
    return token

def get_valid_session(token):
    """Return the emp_id for this token if it exists and hasn't timed out, else None.
    Also cleans up the session row if it's expired."""
    if not token:
        return None
    df = get_as_df("active_sessions")
    if df.empty or 'session_token' not in df.columns:
        return None
    match = df[df['session_token'].astype(str) == str(token)]
    if match.empty:
        return None
    row = match.iloc[0]
    try:
        last_activity = datetime.strptime(str(row['last_activity']), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None
    if (nepal_now() - last_activity).total_seconds() > SESSION_TIMEOUT_MINUTES * 60:
        delete_matching_rows("active_sessions", "session_token", token)
        return None
    return str(row['emp_id'])

def touch_session(token):
    """Reset the idle timer for this session — called on every active page load."""
    if token:
        execute_query("active_sessions", "update", row_id=token, update_dict={"last_activity": nepal_now().strftime("%Y-%m-%d %H:%M:%S")})

def end_session(token):
    if token:
        delete_matching_rows("active_sessions", "session_token", token)

# --- APP UI ---
st.title("⏰ TPMS")

# Only reveal the Manager portal option when accessed via ?role=manager.
# Employees using the plain link never see that a Manager portal exists.
manager_link_used = st.query_params.get("role") == "manager"

if manager_link_used:
    role = st.sidebar.selectbox("Select Portal", ["Employee", "Manager"])
else:
    role = "Employee"

if role == "Employee":
    st.subheader("Employee")
    emp_df = get_as_df("employees")
    
    if emp_df.empty or 'name' not in emp_df.columns:
        st.warning("The 'employees' sheet has no records or is missing required columns.")
    else:
        if "emp_logged_in" not in st.session_state:
            st.session_state.emp_logged_in = False
            st.session_state.current_emp = None
            st.session_state.session_token = None

        # Auto-restore login from the URL's session token, so a page refresh
        # (or a mobile browser reconnecting) doesn't require re-entering the PIN.
        if not st.session_state.emp_logged_in:
            url_token = st.query_params.get("session")
            if url_token:
                restored_emp_id = get_valid_session(url_token)
                if restored_emp_id:
                    match = emp_df[emp_df['emp_id'].astype(str) == str(restored_emp_id)]
                    if not match.empty:
                        st.session_state.emp_logged_in = True
                        st.session_state.current_emp = match.iloc[0].to_dict()
                        st.session_state.session_token = url_token
                        touch_session(url_token)

        if not st.session_state.emp_logged_in:
            st.markdown("### 🔐 Employee Login")
            selected_name = st.selectbox("Select Your Name", emp_df['name'].tolist())
            entered_pin = st.text_input("Enter 4-Digit PIN", type="password", autocomplete="off")
            
            if st.button("Login", type="primary"):
                emp_row = emp_df[emp_df['name'] == selected_name].iloc[0]
                if str(emp_row['pin']) == str(entered_pin):
                    token = create_session(emp_row['emp_id'])
                    st.session_state.emp_logged_in = True
                    st.session_state.current_emp = emp_row.to_dict()
                    st.session_state.session_token = token
                    st.query_params["session"] = token
                    st.rerun()
                else:
                    st.error("Incorrect PIN. Please try again.")
        else:
            emp = st.session_state.current_emp
            touch_session(st.session_state.get("session_token"))
            st.success(f"Welcome, **{emp['name']}**! 👋")
            
            top_c1, top_c2 = st.columns([3, 1])
            with top_c1:
                st.write(f"ID: `{emp['emp_id']}` | Shift: **{emp.get('shift_name', 'Standard')}** | Rate: **NPR {emp.get('hourly_rate', 170)}/hr**")
            with top_c2:
                if st.button("Logout"):
                    end_session(st.session_state.get("session_token"))
                    st.session_state.emp_logged_in = False
                    st.session_state.current_emp = None
                    st.session_state.session_token = None
                    if "session" in st.query_params:
                        del st.query_params["session"]
                    st.rerun()
            
            st.divider()

            emp_tab1, emp_tab2 = st.tabs(["⏱️ Timeclock", "📊 History"])

            with emp_tab1:
                punches_df = get_as_df("time_punches")
                active_punch = None

                if not punches_df.empty and 'emp_id' in punches_df.columns:
                    emp_punches = punches_df[punches_df['emp_id'].astype(str) == str(emp['emp_id'])]
                    open_punches = emp_punches[emp_punches['clock_out'].isna() | (emp_punches['clock_out'].astype(str).str.strip() == "")]
                    if not open_punches.empty:
                        active_punch = open_punches.iloc[-1]

                if active_punch is None:
                    st.info("Status: **Currently Clocked Out** ⚪")
                    
                    if st.button("🟢 Clock In", type="primary", use_container_width=True):
                        if st.session_state.get("clock_in_in_progress"):
                            st.warning("Already processing your clock-in, please wait...")
                        else:
                            st.session_state["clock_in_in_progress"] = True
                            try:
                                now_dt = nepal_now()
                                now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                                actual_shift = determine_shift(now_dt)
                                punch_id = f"P_{int(now_dt.timestamp())}"
                                new_row = [punch_id, emp['emp_id'], actual_shift, now_str, "", 0, 0, 0, 0, "Pending", "", "No"]
                                execute_query("time_punches", "insert", data_row=new_row)
                                st.success(f"Successfully clocked in for {actual_shift} at {now_str} (Nepal time)!")
                                st.rerun()
                            finally:
                                st.session_state["clock_in_in_progress"] = False
                else:
                    clock_in_time = active_punch['clock_in']
                    st.success(f"Status: **Currently Clocked In** 🟢\n\n* Shift: **{active_punch['shift_name']}**\n* Started at: **{clock_in_time}**")

                    # --- LOG A SALE (builds up Cash In live during the shift) ---
                    st.divider()
                    st.markdown("### 💵 Log a Sale")
                    st.caption("Add each payment as you receive it. These add up automatically into Cash In at clock-out — no need to total them yourself.")

                    shift_txn_df = get_as_df("shift_transactions")
                    my_txns = pd.DataFrame()
                    if not shift_txn_df.empty and 'punch_id' in shift_txn_df.columns:
                        my_txns = shift_txn_df[shift_txn_df['punch_id'].astype(str) == str(active_punch['punch_id'])]
                    current_shift_total = my_txns['amount'].apply(safe_float).sum() if not my_txns.empty else 0.0

                    st.metric("Cash In So Far This Shift", f"${current_shift_total:,.2f}")

                    with st.form("add_sale_form", clear_on_submit=True):
                        sale_amount = st.number_input("Sale Amount (USD)", min_value=0.01, value=None, step=1.0, placeholder="Enter amount")
                        add_sale_submitted = st.form_submit_button("➕ Add Sale")

                        if add_sale_submitted:
                            if sale_amount is None or sale_amount <= 0:
                                st.error("Enter a valid amount greater than 0.")
                            elif st.session_state.get("add_sale_in_progress"):
                                st.warning("Already adding, please wait...")
                            else:
                                st.session_state["add_sale_in_progress"] = True
                                try:
                                    txn_time = nepal_now()
                                    txn_id = f"T_{int(txn_time.timestamp() * 1000)}"
                                    new_txn = [txn_id, active_punch['punch_id'], emp['emp_id'], float(sale_amount), txn_time.strftime("%Y-%m-%d %H:%M:%S")]
                                    execute_query("shift_transactions", "insert", data_row=new_txn)
                                    st.success(f"Added ${sale_amount:,.2f}")
                                    st.rerun()
                                finally:
                                    st.session_state["add_sale_in_progress"] = False

                    if not my_txns.empty:
                        with st.expander(f"View {len(my_txns)} entries logged this shift"):
                            st.caption("Made a mistake? Tap 🗑️ next to the wrong entry to remove it.")
                            for _, txn_row in my_txns.sort_values('timestamp').iterrows():
                                col_time, col_amt, col_del = st.columns([3, 2, 1])
                                col_time.write(txn_row['timestamp'])
                                col_amt.write(f"${safe_float(txn_row['amount']):,.2f}")
                                if col_del.button("🗑️", key=f"del_txn_{txn_row['txn_id']}"):
                                    delete_matching_rows("shift_transactions", "txn_id", txn_row['txn_id'])
                                    st.success("Entry removed.")
                                    st.rerun()

                    # --- END OF SHIFT REPORT ---
                    st.divider()
                    st.markdown("### 📝 End of Shift Daily Report")
                    st.caption("Cash In is calculated automatically from what you logged above. Fill in Cash Out and Bonus below — enter 0 if there's nothing to report.")
                    st.info(f"**Cash In (automatic): ${current_shift_total:,.2f}**")
                    
                    with st.form("daily_report_form"):
                        cash_out = st.number_input("Cash Out (USD)", min_value=0.0, value=None, step=1.0, placeholder="Enter amount, or 0")
                        bonus = st.number_input("Customer Bonus / Tips (USD)", min_value=0.0, value=None, step=1.0, placeholder="Enter amount, or 0")
                        
                        submitted = st.form_submit_button("🔴 Submit Report & Clock Out", type="primary", use_container_width=True)
                        
                        if submitted:
                            if cash_out is None or bonus is None:
                                st.error("Please fill in both fields before clocking out. Enter 0 if there's nothing to report.")
                            elif st.session_state.get("clock_out_in_progress"):
                                st.warning("Already processing your clock-out, please wait...")
                            else:
                                st.session_state["clock_out_in_progress"] = True
                                try:
                                    now = nepal_now()
                                    clock_out_str = now.strftime("%Y-%m-%d %H:%M:%S")
                                    try:
                                        t_in = datetime.strptime(str(clock_in_time), "%Y-%m-%d %H:%M:%S")
                                        diff_hours = round((now - t_in).total_seconds() / 3600.0, 2)
                                    except Exception:
                                        diff_hours = 0.0

                                    # Recompute the final total fresh, in case a sale was added moments ago
                                    fresh_txn_df = get_as_df("shift_transactions")
                                    if not fresh_txn_df.empty and 'punch_id' in fresh_txn_df.columns:
                                        final_cash_in = fresh_txn_df[fresh_txn_df['punch_id'].astype(str) == str(active_punch['punch_id'])]['amount'].apply(safe_float).sum()
                                    else:
                                        final_cash_in = 0.0

                                    execute_query(
                                        "time_punches", "update", row_id=active_punch['punch_id'], 
                                        update_dict={"clock_out": clock_out_str, "total_hours": diff_hours, "cash_in": float(final_cash_in), "cash_out": float(cash_out), "bonus": float(bonus)}
                                    )
                                    # Wipe this shift's individual sale entries now that they're summarized above
                                    delete_matching_rows("shift_transactions", "punch_id", active_punch['punch_id'])

                                    st.success(f"Clocked out successfully at {clock_out_str}! Total hours: {diff_hours} hrs. Cash In: ${final_cash_in:,.2f}")
                                    st.rerun()
                                finally:
                                    st.session_state["clock_out_in_progress"] = False

            with emp_tab2:
                st.markdown("### 📊 History")
                punches_df = get_as_df("time_punches")
                if punches_df.empty or 'emp_id' not in punches_df.columns:
                    st.info("No work history available yet.")
                else:
                    emp_history = punches_df[punches_df['emp_id'].astype(str) == str(emp['emp_id'])].copy()
                    emp_history = emp_history[emp_history['clock_out'].notna() & (emp_history['clock_out'].astype(str).str.strip() != "")]
                    
                    if emp_history.empty:
                        st.info("No completed shifts found yet.")
                    else:
                        emp_history['clock_in_dt'] = pd.to_datetime(emp_history['clock_in'], errors='coerce')
                        emp_history['year_str'] = emp_history['clock_in_dt'].dt.year.fillna(2026).astype(int).astype(str)
                        emp_history['month_str'] = emp_history['clock_in_dt'].dt.strftime('%B')
                        
                        def calc_punctuality(row):
                            shift_name = str(row.get('shift_name', '')).lower()
                            dt = row.get('clock_in_dt')
                            if pd.isna(dt): return "On Time"
                            target_hour = 6
                            if 'evening' in shift_name: target_hour = 14
                            elif 'night' in shift_name: target_hour = 22
                            
                            target_mins = target_hour * 60
                            actual_mins = dt.hour * 60 + dt.minute
                            diff_mins = actual_mins - target_mins
                            if diff_mins > 5: return f"Late ({diff_mins} mins late)"
                            elif diff_mins < 0: return f"On Time ({abs(diff_mins)} mins early)"
                            else: return "On Time"

                        emp_history['Punctuality'] = emp_history.apply(calc_punctuality, axis=1)

                        def format_status(val):
                            v = str(val).strip().capitalize()
                            if v == "Approved": return "✔"
                            elif v == "Declined": return "✖"
                            else: return "-"
                        
                        if 'approval_status' in emp_history.columns:
                            emp_history['approval_status'] = emp_history['approval_status'].apply(format_status)

                        s_col1, s_col2 = st.columns(2)
                        months_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
                        current_month_name = nepal_now().strftime('%B')
                        default_m_idx = months_list.index(current_month_name) if current_month_name in months_list else 0
                        years_list = [str(y) for y in range(2026, 2036)]
                        current_year_str = str(nepal_now().year)
                        default_y_idx = years_list.index(current_year_str) if current_year_str in years_list else 0
                        
                        with s_col1: selected_month = st.selectbox("Select Month", months_list, index=default_m_idx)
                        with s_col2: selected_year = st.selectbox("Select Year", years_list, index=default_y_idx)
                        
                        filtered_df = emp_history[(emp_history['year_str'] == selected_year) & (emp_history['month_str'] == selected_month)]
                        total_hours = filtered_df['total_hours'].apply(safe_float).sum()
                        hourly_rate = safe_float(emp.get('hourly_rate', 170) or 170)
                        total_salary = total_hours * hourly_rate
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Total Hours", f"{total_hours:.2f} hrs")
                        m2.metric("Hourly Rate", f"NPR {hourly_rate:,.2f}/hr")
                        m3.metric("Total Monthly Salary", f"NPR {total_salary:,.2f}")
                        
                        st.divider()
                        st.markdown("#### Daily Reports")
                        display_cols = ['shift_name', 'clock_in', 'clock_out', 'total_hours', 'cash_in', 'cash_out', 'bonus', 'Punctuality', 'approval_status']
                        available_display_cols = [c for c in display_cols if c in filtered_df.columns]
                        
                        def highlight_table(df):
                            style_df = pd.DataFrame('', index=df.index, columns=df.columns)
                            for idx, row in df.iterrows():
                                try:
                                    if 'cash_in' in df.columns and 'cash_out' in df.columns:
                                        if safe_float(row['cash_in']) < safe_float(row['cash_out']):
                                            style_df.loc[idx, 'cash_out'] = 'color: #cc0000; font-weight: bold;'
                                except: pass
                                if 'Punctuality' in df.columns:
                                    if 'Late' in str(row['Punctuality']):
                                        style_df.loc[idx, 'Punctuality'] = 'color: #cc0000; font-weight: bold;'
                            return style_df

                        styled_df = filtered_df[available_display_cols].style.apply(highlight_table, axis=None)
                        st.dataframe(styled_df, use_container_width=True)

elif role == "Manager":
    st.subheader("Manager Portal")
    
    if "manager_logged_in" not in st.session_state:
        st.session_state.manager_logged_in = False

    if not st.session_state.manager_logged_in:
        st.markdown("### 🔐 Manager Login")
        mgr_user = st.text_input("Username")
        mgr_pass = st.text_input("Password", type="password", autocomplete="off")
        
        if st.button("Login", type="primary"):
            stored_mgr_pw = get_setting("manager_password", "Money@100")
            if mgr_user == "admin" and mgr_pass == stored_mgr_pw:
                st.session_state.manager_logged_in = True
                st.rerun()
            else:
                st.error("Incorrect username or password. Please try again.")
    else:
        top_c1, top_c2 = st.columns([3, 1])
        with top_c1: st.success("Logged in as: **Manager**")
        with top_c2:
            if st.button("Logout"):
                st.session_state.manager_logged_in = False
                st.rerun()
        
        st.divider()
        
        mgr_tab1, mgr_tab2, mgr_tab3, mgr_tab4, mgr_tab5 = st.tabs(["👥 Employees & Shifts", "💰 Financial Dashboard", "⏱️ Time Punches", "📊 Expenses", "⚙️ Settings"])
        
        with mgr_tab1:
            st.markdown("### Employee & Shift Time Slot Management")
            emp_df = get_as_df("employees")
            st.dataframe(emp_df, use_container_width=True)
            
            st.markdown("#### Add New Employee & Assign Shift Slot")
            with st.form("add_employee_form"):
                new_emp_name = st.text_input("Employee Name")
                new_emp_rate = st.number_input("Hourly Rate (NPR)", min_value=0.0, value=170.0, step=10.0)
                new_emp_pin = st.text_input("4-Digit PIN", max_chars=4, type="password", autocomplete="off")
                shift_choices = [
                    "Morning Shift (06:00 - 14:00)",
                    "Evening Shift (14:00 - 22:00)",
                    "Night Shift (22:00 - 06:00)"
                ]
                assigned_shift = st.selectbox("Assign 8-Hour Time Slot", shift_choices)
                add_emp_btn = st.form_submit_button("Add Employee", type="primary")
                
                if add_emp_btn and new_emp_name and new_emp_pin:
                    emp_id = f"E_{int(nepal_now().timestamp())}"
                    execute_query("employees", "insert", data_row=[emp_id, new_emp_name, new_emp_rate, new_emp_pin, assigned_shift])
                    st.success(f"Employee '{new_emp_name}' added successfully with shift '{assigned_shift}'!")
                    st.rerun()

            st.divider()
            st.markdown("#### ✏️ Edit Employee (Shift / Rate)")
            st.caption("Use this to fix an employee's assigned shift or rate — e.g. if someone is clocking in for Night Shift but was set up as Morning Shift.")

            if emp_df.empty or 'name' not in emp_df.columns:
                st.info("No employees found.")
            else:
                with st.form("edit_employee_form"):
                    edit_emp_name = st.selectbox("Select Employee", emp_df['name'].tolist(), key="edit_emp_select")
                    edit_emp_row = emp_df[emp_df['name'] == edit_emp_name].iloc[0]

                    current_rate = safe_float(edit_emp_row.get('hourly_rate', 170))
                    current_shift = str(edit_emp_row.get('shift_name', 'Morning Shift (06:00 - 14:00)'))
                    shift_choices_edit = [
                        "Morning Shift (06:00 - 14:00)",
                        "Evening Shift (14:00 - 22:00)",
                        "Night Shift (22:00 - 06:00)"
                    ]
                    default_idx = shift_choices_edit.index(current_shift) if current_shift in shift_choices_edit else 0

                    new_shift = st.selectbox("Assigned Shift", shift_choices_edit, index=default_idx, key="edit_emp_shift")
                    new_rate = st.number_input("Hourly Rate (NPR)", min_value=0.0, value=current_rate, step=10.0, key="edit_emp_rate")
                    edit_submit = st.form_submit_button("💾 Save Changes", type="primary")

                    if edit_submit:
                        execute_query(
                            "employees", "update", row_id=edit_emp_row['emp_id'],
                            update_dict={"shift_name": new_shift, "hourly_rate": new_rate}
                        )
                        st.success(f"Updated {edit_emp_name}: shift set to '{new_shift}', rate set to NPR {new_rate:,.2f}/hr.")
                        st.rerun()

        with mgr_tab2:
            st.markdown("### 💰 Financial Dashboard")
            EXCHANGE_RATE = 133.0

            f_col1, f_col2 = st.columns(2)
            with f_col1: view_mode = st.selectbox("View Slicer", ["All-Time", "Monthly", "Daily"], key="fd_view_slicer")
            
            punches_df = get_as_df("time_punches")
            expenses_df = get_as_df("expenses")
            emps_df = get_as_df("employees")
            
            if not punches_df.empty and 'clock_in' in punches_df.columns:
                punches_df['dt'] = pd.to_datetime(punches_df['clock_in'], errors='coerce')
                punches_df['date_str'] = punches_df['dt'].dt.strftime('%Y-%m-%d')
                punches_df['month_year'] = punches_df['dt'].dt.strftime('%B %Y')
            
            if not expenses_df.empty and 'expense_date' in expenses_df.columns:
                expenses_df['dt'] = pd.to_datetime(expenses_df['expense_date'], errors='coerce')
                expenses_df['date_str'] = expenses_df['dt'].dt.strftime('%Y-%m-%d')
                expenses_df['month_year'] = expenses_df['dt'].dt.strftime('%B %Y')

            filtered_punches = punches_df.copy()
            filtered_expenses = expenses_df.copy()

            if view_mode == "Monthly" and not punches_df.empty:
                available_months = sorted(punches_df['month_year'].dropna().unique(), reverse=True)
                sel_month = st.selectbox("Select Month", available_months if available_months else [nepal_now().strftime('%B %Y')], key="fd_sel_month")
                filtered_punches = punches_df[punches_df['month_year'] == sel_month]
                if not expenses_df.empty: filtered_expenses = expenses_df[expenses_df['month_year'] == sel_month]
            elif view_mode == "Daily" and not punches_df.empty:
                available_dates = sorted(punches_df['date_str'].dropna().unique(), reverse=True)
                sel_date = st.selectbox("Select Date", available_dates if available_dates else [str(nepal_today())], key="fd_sel_date")
                filtered_punches = punches_df[punches_df['date_str'] == sel_date]
                if not expenses_df.empty: filtered_expenses = expenses_df[expenses_df['date_str'] == sel_date]

            total_cash_in = sum(safe_float(x) for x in filtered_punches['cash_in']) if not filtered_punches.empty and 'cash_in' in filtered_punches.columns else 0.0
            total_cash_out = sum(safe_float(x) for x in filtered_punches['cash_out']) if not filtered_punches.empty and 'cash_out' in filtered_punches.columns else 0.0
            total_bonus = sum(safe_float(x) for x in filtered_punches['bonus']) if not filtered_punches.empty and 'bonus' in filtered_punches.columns else 0.0
            
            revenue = total_cash_in - total_cash_out - (1.0 / 6.0) * total_bonus

            salary_expense_npr = 0.0
            if not filtered_punches.empty and not emps_df.empty:
                rate_map = dict(zip(emps_df['emp_id'].astype(str), emps_df['hourly_rate'].apply(safe_float)))
                for _, row in filtered_punches.iterrows():
                    e_id = str(row.get('emp_id', ''))
                    hrs = safe_float(row.get('total_hours', 0))
                    rate = rate_map.get(e_id, 170.0)
                    salary_expense_npr += hrs * rate
            salary_expense = salary_expense_npr / EXCHANGE_RATE

            op_expense = 0.0
            other_expense = 0.0
            if not filtered_expenses.empty and 'amount' in filtered_expenses.columns:
                for _, row in filtered_expenses.iterrows():
                    cat = str(row.get('category', '')).lower()
                    amt = safe_float(row.get('amount', 0))
                    if 'other' in cat: other_expense += amt
                    else: op_expense += amt

            total_expenses = salary_expense + op_expense + other_expense
            profit = revenue - total_expenses

            kpi1, kpi2 = st.columns(2)
            kpi1.metric("Revenue", f"${revenue:,.2f}")
            kpi2.metric("Net Profit", f"${profit:,.2f}", delta=f"{(profit/revenue*100):.1f}% margin" if revenue > 0 else "0.0%")
            
            st.divider()
            st.markdown("#### Financial Summary")
            fin_summary_df = pd.DataFrame({
                "Financial Metric": ["Revenue", "Salary Expense", "Operating Expense", "Other Expense", "Total Expenses", "Net Profit"],
                "Amount (USD)": [f"${revenue:,.2f}", f"${salary_expense:,.2f}", f"${op_expense:,.2f}", f"${other_expense:,.2f}", f"${total_expenses:,.2f}", f"${profit:,.2f}"]
            })
            st.dataframe(fin_summary_df, use_container_width=True)

            st.divider()
            st.markdown("#### Financial Trend Visualization")
            if not filtered_punches.empty and 'date_str' in filtered_punches.columns:
                chart_df = filtered_punches.groupby('date_str').agg({
                    'cash_in': lambda x: sum(safe_float(v) for v in x),
                    'cash_out': lambda x: sum(safe_float(v) for v in x),
                    'bonus': lambda x: sum(safe_float(v) for v in x)
                }).reset_index()
                chart_df['Revenue'] = chart_df['cash_in'] - chart_df['cash_out'] - (1.0 / 6.0) * chart_df['bonus']
                chart_df['Profit'] = chart_df['Revenue']
                
                fig = px.line(chart_df, x='date_str', y=['cash_in', 'cash_out', 'Revenue', 'Profit'],
                              labels={'value': 'Amount (USD)', 'date_str': 'Date', 'variable': 'Metric'},
                              title="Cash In, Cash Out, Revenue, and Profit Trends")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Insufficient data for trend visualization.")

        with mgr_tab3:
            st.markdown("### ⏱️ Time Punches & Reports")
            
            punches_df = get_as_df("time_punches")
            emps_df = get_as_df("employees")
            
            if punches_df.empty:
                st.info("No time punches recorded yet.")
            else:
                emp_name_map = {}
                if not emps_df.empty and 'emp_id' in emps_df.columns and 'name' in emps_df.columns:
                    emp_name_map = dict(zip(emps_df['emp_id'].astype(str), emps_df['name']))
                
                punches_df['employee_name'] = punches_df['emp_id'].astype(str).map(emp_name_map).fillna(punches_df['emp_id'])
                
                if 'clock_in' in punches_df.columns:
                    punches_df['dt'] = pd.to_datetime(punches_df['clock_in'], errors='coerce')
                    punches_df['date_str'] = punches_df['dt'].dt.strftime('%Y-%m-%d')
                    punches_df['month_year'] = punches_df['dt'].dt.strftime('%B %Y')
                
                # --- PENDING APPROVALS & EDITABLE QUEUE ---
                pending_df = punches_df[punches_df['approval_status'].astype(str).str.strip().str.capitalize() == "Pending"].copy()
                if not pending_df.empty:
                    st.markdown("#### ⏳ Pending Approvals & Editable Queue")
                    st.info("💡 Edit any incorrect times, hours, or cash entries directly in the table cells below. Check the box under 'Approve?' for the rows you want to approve, then click the button below.")
                    
                    pending_df['Approve?'] = False
                    
                    edit_cols = ['punch_id', 'employee_name', 'shift_name', 'clock_in', 'clock_out', 'total_hours', 'cash_in', 'cash_out', 'bonus', 'Approve?']
                    avail_edit_cols = [c for c in edit_cols if c in pending_df.columns]
                    
                    edited_pending = st.data_editor(
                        pending_df[avail_edit_cols],
                        key="pending_table_editor",
                        use_container_width=True,
                        disabled=['punch_id', 'employee_name'],
                        column_config={
                            "shift_name": st.column_config.SelectboxColumn(
                                "Shift",
                                options=[
                                    "Morning Shift (06:00 - 14:00)",
                                    "Evening Shift (14:00 - 22:00)",
                                    "Night Shift (22:00 - 06:00)"
                                ],
                                required=True
                            )
                        }
                    )
                    
                    col_p1, col_p2 = st.columns(2)
                    if col_p1.button("💾 Save Edits Only", key="btn_save_edits"):
                        ws = gsheet.worksheet("time_punches")
                        headers = ws.row_values(1)
                        for _, row in edited_pending.iterrows():
                            p_id = row['punch_id']
                            cell = ws.find(str(p_id))
                            if cell:
                                row_num = cell.row
                                for field in ['clock_in', 'clock_out', 'total_hours', 'cash_in', 'cash_out', 'bonus']:
                                    if field in headers:
                                        col_num = headers.index(field) + 1
                                        ws.update_cell(row_num, col_num, row[field])
                        st.cache_data.clear()
                        st.success("Changes saved successfully!")
                        st.rerun()
                        
                    if col_p2.button("✔ Save & Approve Selected", type="primary", key="btn_save_approve"):
                        ws = gsheet.worksheet("time_punches")
                        headers = ws.row_values(1)
                        approved_count = 0
                        for _, row in edited_pending.iterrows():
                            p_id = row['punch_id']
                            cell = ws.find(str(p_id))
                            if cell:
                                row_num = cell.row
                                for field in ['clock_in', 'clock_out', 'total_hours', 'cash_in', 'cash_out', 'bonus']:
                                    if field in headers:
                                        col_num = headers.index(field) + 1
                                        ws.update_cell(row_num, col_num, row[field])
                                
                                if row.get('Approve?', False):
                                    if 'approval_status' in headers:
                                        col_num = headers.index('approval_status') + 1
                                        ws.update_cell(row_num, col_num, "Approved")
                                        approved_count += 1
                        st.cache_data.clear()
                        st.success(f"Changes saved and {approved_count} shift(s) approved!")
                        st.rerun()
                    st.divider()

                # --- TIME PUNCHES SLICERS ---
                emp_names_list = ["All Employees"] + sorted(list(set(emp_name_map.values()))) if emp_name_map else ["All Employees"]
                
                tp_c1, tp_c2 = st.columns(2)
                with tp_c1:
                    sel_emp_filter = st.selectbox("Filter by Employee", emp_names_list, key="tp_emp_filter")
                with tp_c2:
                    tp_view_mode = st.selectbox("View Slicer", ["All-Time", "Monthly", "Daily"], key="tp_view_slicer")
                
                filtered_table_df = punches_df.copy()
                if sel_emp_filter != "All Employees":
                    filtered_table_df = filtered_table_df[filtered_table_df['employee_name'] == sel_emp_filter]
                
                if tp_view_mode == "Monthly" and not filtered_table_df.empty:
                    available_months = sorted(filtered_table_df['month_year'].dropna().unique().tolist(), reverse=True)
                    sel_month = st.selectbox("Select Month", available_months if available_months else [nepal_now().strftime('%B %Y')], key="tp_sel_month")
                    filtered_table_df = filtered_table_df[filtered_table_df['month_year'] == sel_month]
                elif tp_view_mode == "Daily" and not filtered_table_df.empty:
                    available_dates = sorted(filtered_table_df['date_str'].dropna().unique().tolist(), reverse=True)
                    sel_date = st.selectbox("Select Date", available_dates if available_dates else [str(nepal_today())], key="tp_sel_date")
                    filtered_table_df = filtered_table_df[filtered_table_df['date_str'] == sel_date]

                st.divider()
                st.markdown("#### 📋 Time Sheets & Cash Reports Table")
                
                if filtered_table_df.empty:
                    st.info("No records match the selected filters.")
                else:
                    def fmt_stat(v):
                        v_str = str(v).strip().capitalize()
                        if v_str == "Approved": return "✔"
                        elif v_str == "Declined": return "✖"
                        else: return "-"
                    
                    filtered_table_df['Status'] = filtered_table_df['approval_status'].apply(fmt_stat)
                    
                    display_columns = ['punch_id', 'employee_name', 'shift_name', 'clock_in', 'clock_out', 'total_hours', 'cash_in', 'cash_out', 'bonus', 'Status']
                    avail_cols = [c for c in display_columns if c in filtered_table_df.columns]
                    st.dataframe(filtered_table_df[avail_cols], use_container_width=True)

        with mgr_tab4:
            st.markdown("### Expenses Management")
            
            EXCHANGE_RATE = 133.0
            all_punches = get_as_df("time_punches")
            all_expenses = get_as_df("expenses")
            all_emps = get_as_df("employees")

            if not all_punches.empty and 'clock_in' in all_punches.columns:
                all_punches['dt'] = pd.to_datetime(all_punches['clock_in'], errors='coerce')
                all_punches['month_year'] = all_punches['dt'].dt.strftime('%B %Y')
                all_punches['date_str'] = all_punches['dt'].dt.strftime('%Y-%m-%d')
            
            if not all_expenses.empty and 'expense_date' in all_expenses.columns:
                all_expenses['dt'] = pd.to_datetime(all_expenses['expense_date'], errors='coerce')
                all_expenses['month_year'] = all_expenses['dt'].dt.strftime('%B %Y')
                all_expenses['date_str'] = all_expenses['dt'].dt.strftime('%Y-%m-%d')

            # Initialize variables safely
            salary_exp_usd = 0.0
            op_exp_usd = 0.0
            oth_exp_usd = 0.0
            sel_exp_month = nepal_now().strftime('%B %Y')

            # --- EXPENSES TAB VIEW SLICER ---
            exp_view_mode = st.selectbox("View Slicer", ["All-Time", "Monthly", "Daily"], key="exp_view_slicer")

            filtered_exp_df = all_expenses.copy()
            filtered_punches_exp = all_punches.copy()

            if exp_view_mode == "Monthly" and not all_expenses.empty:
                available_exp_months = sorted(list(set(
                    (all_expenses['month_year'].dropna().tolist() if 'month_year' in all_expenses.columns else []) +
                    (all_punches['month_year'].dropna().tolist() if not all_punches.empty and 'month_year' in all_punches.columns else [])
                )), reverse=True)
                sel_exp_month = st.selectbox("Select Month", available_exp_months if available_exp_months else [nepal_now().strftime('%B %Y')], key="exp_sel_month")
                
                filtered_exp_df = all_expenses[all_expenses['month_year'] == sel_exp_month] if not all_expenses.empty and 'month_year' in all_expenses.columns else pd.DataFrame()
                filtered_punches_exp = all_punches[all_punches['month_year'] == sel_exp_month] if not all_punches.empty and 'month_year' in all_punches.columns else pd.DataFrame()

            elif exp_view_mode == "Daily" and not all_expenses.empty:
                available_exp_dates = sorted(list(set(
                    (all_expenses['date_str'].dropna().tolist() if 'date_str' in all_expenses.columns else []) +
                    (all_punches['date_str'].dropna().tolist() if not all_punches.empty and 'date_str' in all_punches.columns else [])
                )), reverse=True)
                sel_exp_date = st.selectbox("Select Date", available_exp_dates if available_exp_dates else [str(nepal_today())], key="exp_sel_date")
                
                filtered_exp_df = all_expenses[all_expenses['date_str'] == sel_exp_date] if not all_expenses.empty and 'date_str' in all_expenses.columns else pd.DataFrame()
                filtered_punches_exp = all_punches[all_punches['date_str'] == sel_exp_date] if not all_punches.empty and 'date_str' in all_punches.columns else pd.DataFrame()

            st.divider()

            # --- MONTHLY EMPLOYEE PAYROLL SUMMARY TABLE (TOP CORNER OF EXPENSES TAB) ---
            st.markdown(f"#### 👥 Employee Monthly Hours & Salary Summary ({exp_view_mode}: {sel_exp_month if exp_view_mode=='Monthly' else ('All-Time' if exp_view_mode=='All-Time' else locals().get('sel_exp_date',''))})")
            if not all_emps.empty:
                emp_summary_list = []
                for _, emp_row in all_emps.iterrows():
                    e_id = str(emp_row['emp_id'])
                    e_name = emp_row['name']
                    e_rate = safe_float(emp_row['hourly_rate'])
                    
                    if not filtered_punches_exp.empty and 'emp_id' in filtered_punches_exp.columns:
                        emp_filtered_p = filtered_punches_exp[filtered_punches_exp['emp_id'].astype(str) == e_id]
                        e_hours = emp_filtered_p['total_hours'].apply(safe_float).sum()
                    else:
                        e_hours = 0.0
                    e_salary = e_hours * e_rate
                    
                    emp_summary_list.append({
                        "Employee Name": e_name,
                        "Total Hours": round(e_hours, 2),
                        "Hourly Rate (NPR)": f"{e_rate:,.2f}",
                        "Total Salary (NPR)": f"{e_salary:,.2f}"
                    })
                emp_summary_df = pd.DataFrame(emp_summary_list)
                st.dataframe(emp_summary_df, use_container_width=True)
            else:
                st.info("No employee records found.")

            st.divider()

            # Calculate Salary Expense for filtered period
            sal_exp_npr = 0.0
            if not filtered_punches_exp.empty and not all_emps.empty:
                rate_map = dict(zip(all_emps['emp_id'].astype(str), all_emps['hourly_rate'].apply(safe_float)))
                for _, row in filtered_punches_exp.iterrows():
                    e_id = str(row.get('emp_id', ''))
                    hrs = safe_float(row.get('total_hours', 0))
                    rate = rate_map.get(e_id, 170.0)
                    sal_exp_npr += hrs * rate
            salary_expense_usd = sal_exp_npr / EXCHANGE_RATE

            if not filtered_exp_df.empty and 'amount' in filtered_exp_df.columns:
                for _, row in filtered_exp_df.iterrows():
                    cat = str(row.get('category', '')).lower()
                    amt = safe_float(row.get('amount', 0))
                    if 'other' in cat: oth_exp_usd += amt
                    else: op_exp_usd += amt

            st.markdown(f"#### Expense Breakdown ({exp_view_mode} View)")
            pie_df = pd.DataFrame({
                "Category": ["Salary Expense", "Operating Expenses", "Other Expenses"],
                "Amount (USD)": [salary_exp_usd, op_exp_usd, oth_exp_usd]
            })
            if pie_df["Amount (USD)"].sum() > 0:
                fig_pie = px.pie(pie_df, names="Category", values="Amount (USD)", title="Expense Distribution by Category", hole=0.3)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("No expense data available for the selected period to display the pie chart.")

            st.divider()

            st.markdown("#### Month-over-Month Comparison (This Month vs. Previous Month)")
            try:
                ref_month = sel_exp_month if exp_view_mode == "Monthly" and 'sel_exp_month' in locals() else nepal_now().strftime('%B %Y')
                dt_curr = datetime.strptime(ref_month, '%B %Y')
                prev_month_dt = dt_curr - timedelta(days=28)
                prev_month_str = prev_month_dt.strftime('%B %Y')
            except:
                ref_month = nepal_now().strftime('%B %Y')
                prev_month_str = ""

            def get_month_totals(m_str):
                p_m = all_punches[all_punches['month_year'] == m_str] if not all_punches.empty and 'month_year' in all_punches.columns else pd.DataFrame()
                e_m = all_expenses[all_expenses['month_year'] == m_str] if not all_expenses.empty and 'month_year' in all_expenses.columns else pd.DataFrame()
                
                s_npr = 0.0
                if not p_m.empty and not all_emps.empty:
                    rate_map = dict(zip(all_emps['emp_id'].astype(str), all_emps['hourly_rate'].apply(safe_float)))
                    for _, row in p_m.iterrows():
                        e_id = str(row.get('emp_id', ''))
                        hrs = safe_float(row.get('total_hours', 0))
                        rate = rate_map.get(e_id, 170.0)
                        s_npr += hrs * rate
                s_usd = s_npr / EXCHANGE_RATE

                o_usd = 0.0
                ot_usd = 0.0
                if not e_m.empty and 'amount' in e_m.columns:
                    for _, row in e_m.iterrows():
                        cat = str(row.get('category', '')).lower()
                        amt = safe_float(row.get('amount', 0))
                        if 'other' in cat: ot_usd += amt
                        else: o_usd += amt
                return s_usd, o_usd, ot_usd

            curr_sal, curr_op, curr_oth = get_month_totals(ref_month)
            prev_sal, prev_op, prev_oth = get_month_totals(prev_month_str)

            comp_df = pd.DataFrame({
                "Category": ["Salary Expense", "Operating Expenses", "Other Expenses", "Salary Expense", "Operating Expenses", "Other Expenses"],
                "Month": [ref_month, ref_month, ref_month, prev_month_str, prev_month_str, prev_month_str],
                "Amount (USD)": [curr_sal, curr_op, curr_oth, prev_sal, prev_op, prev_oth]
            })

            fig_bar = px.bar(comp_df, x="Category", y="Amount (USD)", color="Month", barmode="group",
                             title=f"Comparison: {ref_month} vs {prev_month_str}")
            st.plotly_chart(fig_bar, use_container_width=True)

            st.divider()

            st.markdown("#### Add New Expense")
            with st.form("add_expense_form"):
                exp_category = st.selectbox("Expense Category", ["Operating", "Other"])
                exp_vendor = st.text_input("Vendor Name")
                exp_cost = st.number_input("Cost (USD)", min_value=0.0, value=0.0, step=10.0)
                exp_date = st.date_input("Date Settled", value=nepal_today())
                exp_desc = st.text_input("Description / Notes")
                
                submit_exp = st.form_submit_button("Submit Expense", type="primary")
                if submit_exp and exp_vendor:
                    exp_id = f"EXP_{int(nepal_now().timestamp())}"
                    new_exp_row = [exp_id, exp_date.strftime("%Y-%m-%d"), exp_category, exp_vendor, exp_cost, exp_desc]
                    execute_query("expenses", "insert", data_row=new_exp_row)
                    st.success(f"Expense of ${exp_cost:,.2f} to {exp_vendor} added successfully!")
                    st.rerun()

            st.divider()
            st.markdown(f"#### Logged Expenses Table ({exp_view_mode} View)")
            if filtered_exp_df.empty:
                st.info("No expenses logged for this period.")
            else:
                st.dataframe(filtered_exp_df, use_container_width=True)

        with mgr_tab5:
            st.markdown("### ⚙️ Settings")

            st.markdown("#### 🔑 Change Manager Password")
            with st.form("change_mgr_pw_form"):
                current_pw = st.text_input("Current Password", type="password", autocomplete="off")
                new_pw = st.text_input("New Password", type="password", autocomplete="off")
                confirm_pw = st.text_input("Confirm New Password", type="password", autocomplete="off")
                submit_pw = st.form_submit_button("Update Password", type="primary")

                if submit_pw:
                    stored_pw = get_setting("manager_password", "Money@100")
                    if current_pw != stored_pw:
                        st.error("Current password is incorrect.")
                    elif not new_pw:
                        st.error("New password cannot be empty.")
                    elif new_pw != confirm_pw:
                        st.error("New password and confirmation do not match.")
                    else:
                        set_setting("manager_password", new_pw)
                        st.success("Manager password updated successfully! Use it next time you log in.")

            st.divider()

            st.markdown("#### 📥 Add Past Shift (Backfill)")
            st.caption("Use this to manually enter shifts that already happened — e.g. from WhatsApp reports — since they won't come through the live Clock In/Out flow.")

            emp_df_backfill = get_as_df("employees")
            if emp_df_backfill.empty or 'name' not in emp_df_backfill.columns:
                st.info("No employees found.")
            else:
                with st.form("add_past_shift_form"):
                    bf_c1, bf_c2 = st.columns(2)
                    with bf_c1:
                        bf_emp_name = st.selectbox("Employee", emp_df_backfill['name'].tolist(), key="bf_emp")
                        bf_date = st.date_input("Shift Date", value=nepal_today(), key="bf_date")
                        bf_shift = st.selectbox("Shift", [
                            "Morning Shift (06:00 - 14:00)",
                            "Evening Shift (14:00 - 22:00)",
                            "Night Shift (22:00 - 06:00)"
                        ], key="bf_shift")
                    with bf_c2:
                        bf_time_in = st.time_input("Clock In Time", value=time(6, 0), key="bf_in")
                        bf_time_out = st.time_input("Clock Out Time", value=time(14, 0), key="bf_out")

                    bf_c3, bf_c4, bf_c5 = st.columns(3)
                    with bf_c3:
                        bf_cash_in = st.number_input("Cash In (USD)", min_value=0.0, value=0.0, step=1.0, key="bf_cash_in")
                    with bf_c4:
                        bf_cash_out = st.number_input("Cash Out (USD)", min_value=0.0, value=0.0, step=1.0, key="bf_cash_out")
                    with bf_c5:
                        bf_bonus = st.number_input("Bonus (USD)", min_value=0.0, value=0.0, step=1.0, key="bf_bonus")

                    bf_submit = st.form_submit_button("➕ Add Shift Record", type="primary")

                    if bf_submit:
                        bf_dt_in = datetime.combine(bf_date, bf_time_in)
                        bf_dt_out = datetime.combine(bf_date, bf_time_out)
                        if bf_dt_out <= bf_dt_in:
                            bf_dt_out = bf_dt_out + timedelta(days=1)  # handles overnight shifts like Night Shift

                        bf_hours = round((bf_dt_out - bf_dt_in).total_seconds() / 3600.0, 2)
                        bf_emp_row = emp_df_backfill[emp_df_backfill['name'] == bf_emp_name].iloc[0]
                        bf_punch_id = f"P_bf_{int(nepal_now().timestamp())}"

                        new_row = [
                            bf_punch_id,
                            bf_emp_row['emp_id'],
                            bf_shift,
                            bf_dt_in.strftime("%Y-%m-%d %H:%M:%S"),
                            bf_dt_out.strftime("%Y-%m-%d %H:%M:%S"),
                            bf_hours,
                            float(bf_cash_in),
                            float(bf_cash_out),
                            float(bf_bonus),
                            "Approved",
                            "",
                            "No"
                        ]
                        execute_query("time_punches", "insert", data_row=new_row)
                        st.success(f"Added: {bf_emp_name} — {bf_date} — {bf_hours} hrs")
                        st.rerun()

            st.divider()
            st.markdown("#### 🔢 Change Employee PIN")
            emp_df_settings = get_as_df("employees")
            if emp_df_settings.empty or 'name' not in emp_df_settings.columns:
                st.info("No employees found.")
            else:
                with st.form("change_emp_pin_form"):
                    sel_emp_for_pin = st.selectbox("Select Employee", emp_df_settings['name'].tolist())
                    new_pin = st.text_input("New 4-Digit PIN", max_chars=4, type="password", autocomplete="off")
                    confirm_pin = st.text_input("Confirm New PIN", max_chars=4, type="password", autocomplete="off")
                    submit_pin = st.form_submit_button("Update PIN", type="primary")

                    if submit_pin:
                        if not new_pin or len(new_pin) != 4 or not new_pin.isdigit():
                            st.error("PIN must be exactly 4 digits.")
                        elif new_pin != confirm_pin:
                            st.error("PINs do not match.")
                        else:
                            emp_row = emp_df_settings[emp_df_settings['name'] == sel_emp_for_pin].iloc[0]
                            execute_query("employees", "update", row_id=emp_row['emp_id'], update_dict={"pin": new_pin})
                            st.success(f"PIN updated for {sel_emp_for_pin}!")
                            st.rerun()
